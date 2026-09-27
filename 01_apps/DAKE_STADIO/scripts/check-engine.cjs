require('./local-test-env.cjs')();
const path = require('node:path');
const fs = require('node:fs');
const root = path.resolve(__dirname, '..');
const evidence = path.join(root, 'evidence');
if (!process.versions.electron) {
  require('esbuild').buildSync({ entryPoints: [path.join(root, 'tests/engine-browser.js')], outfile: path.join(evidence, 'engine-check.bundle.js'), bundle: true, format: 'iife', platform: 'browser' });
  require('esbuild').buildSync({entryPoints:[path.join(root,'src/vector-worker.js')],outfile:path.join(evidence,'vector-worker.js'),bundle:true,format:'iife',platform:'browser',legalComments:'eof'});
  fs.writeFileSync(path.join(evidence, 'engine-check.html'), '<!doctype html><meta charset="utf-8"><title>Engine checks</title><canvas id="editor"></canvas><script src="engine-check.bundle.js"></script>');
  fs.copyFileSync(path.join(root, 'src/raster-worker.js'), path.join(evidence, 'raster-worker.js'));
  const child = require('node:child_process').spawnSync(require('electron'), [__filename, ...process.argv.slice(2)], { cwd: root, stdio: 'inherit', windowsHide: true });
  process.exit(child.status ?? 1);
} else {
  const { app, BrowserWindow } = require('electron');
  app.setPath('userData',path.join(root,'test-output','engine-user-data'));
  app.whenReady().then(async () => {
    const window = new BrowserWindow({ show: false, width: 1000, height: 800, webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false } });
    window.webContents.on('console-message', (event) => { if (/Error|error/i.test(event.message)) console.error(event.message); });
    await window.loadFile(path.join(evidence, 'engine-check.html'));
    try {
      const report = await window.webContents.executeJavaScript('window.runEngineChecks()');
      if (report.finalDocument) { fs.writeFileSync(path.join(evidence, 'engine-roundtrip.dake'), JSON.stringify(report.finalDocument)); delete report.finalDocument; }
      if (report.screenshot) { fs.writeFileSync(path.join(evidence, 'engine-export.png'), Buffer.from(report.screenshot.split(',')[1], 'base64')); delete report.screenshot; }
      if (report.svg) { fs.writeFileSync(path.join(evidence, 'engine-export.svg'), report.svg); delete report.svg; }
      if (report.passed) {
        const coords = await window.webContents.executeJavaScript('window.prepareNativeInputChecks()');
        const delay = () => new Promise((resolve) => setTimeout(resolve, 35));
        const drag = async (from, to) => {
          window.webContents.sendInputEvent({ type:'mouseMove', ...from }); await delay();
          window.webContents.sendInputEvent({ type:'mouseDown', ...from, button:'left', clickCount:1 }); await delay();
          window.webContents.sendInputEvent({ type:'mouseMove', ...to, button:'left', modifiers:['leftButtonDown'] }); await delay();
          window.webContents.sendInputEvent({ type:'mouseUp', ...to, button:'left', clickCount:1 }); await delay();
        };
        await drag(coords.start,coords.end);
        const brushPixel = await window.webContents.executeJavaScript('window.checkNativePixels()');
        const brushPassed = brushPixel[2] === 255 && brushPixel[0] === 0;
        report.results.push({ name:'Electron native mouse events paint raster pixels', passed:brushPassed, pixel:brushPixel });
        await window.webContents.executeJavaScript("window.engine.setTool('eraser')");
        await drag(coords.start,coords.end);
        const eraserPixel = await window.webContents.executeJavaScript('window.checkNativePixels()');
        report.results.push({ name:'Electron native mouse events erase raster alpha', passed:eraserPixel[3] === 0, pixel:eraserPixel });
        const filteredHistory = await window.webContents.executeJavaScript("(async()=>{await window.engine.applyImageAdjustment({brightness:0.1,contrast:0.12});window.engine.setTool('brush');return window.engine._historyIndex;})()");
        await drag(coords.start,coords.end);
        const nativeFiltered = await window.webContents.executeJavaScript('(async()=>{await window.engine._strokeTask;return {history:window.engine._historyIndex,filters:window.engine.selected.filters.length,pixel:window.checkNativePixels()};})()');
        report.results.push({name:'Electron native filtered brush stroke retains filters and one undo step',passed:nativeFiltered.history===filteredHistory+1&&nativeFiltered.filters===2&&nativeFiltered.pixel[2]===255&&nativeFiltered.pixel[3]===255});
        await window.webContents.executeJavaScript('window.engine._busy=true;window.engine._applyTool()');
        await drag(coords.start,coords.end);
        const busySelection = await window.webContents.executeJavaScript('(()=>{const ok=window.engine.selected?.id===window.nativeImageId&&window.engine.layers.length===1;window.engine._busy=false;window.engine._applyTool();return ok;})()');
        report.results.push({name:'Native pointer input while busy preserves selected raster layer',passed:busySelection});
        await window.webContents.executeJavaScript("window.engine.setTool('pen')");
        await drag(coords.pen1,coords.pen1handle); await drag(coords.pen2,coords.pen2handle);
        const nativePath = await window.webContents.executeJavaScript('window.engine.finishPath(false)?.path');
        const penPassed = nativePath?.[1]?.[0] === 'C' && Math.abs(nativePath[1][1]-40)<1 && Math.abs(nativePath[1][2]-180)<1;
        report.results.push({ name:'Electron native drag events create cubic pen path', passed:!!penPassed, path:nativePath });
        const node = await window.webContents.executeJavaScript(`(() => {
          const e=window.engine; e.setTool('node'); e.canvas.renderAll();
          const p=e.selected; p.setCoords(); const box=e.canvas.upperCanvasEl.getBoundingClientRect();
          const c=p.oCoords[Object.keys(p.controls)[0]];
          return {x:Math.round(box.left+c.x),y:Math.round(box.top+c.y),oldX:p.path[0][1],history:e._historyIndex};
        })()`);
        await drag({x:node.x,y:node.y},{x:node.x+26,y:node.y-13});
        const edited = await window.webContents.executeJavaScript('({x:window.engine.selected.path[0][1],left:window.engine.selected.left,history:window.engine._historyIndex})');
        report.results.push({ name:'Electron native node drag edits path with one undo step', passed:Number.isFinite(edited.left)&&Math.abs(edited.x-node.oldX)>5&&edited.history===node.history+1 });
        report.passed = report.results.every((item) => item.passed);
        const foundation=await window.webContents.executeJavaScript('window.runFoundationChecks()');
        report.results.push(...foundation.results);report.passed=report.results.every(item=>item.passed);
        const fixtureDefinitions=[['arial.ttf','ArialMT',400,false],['arialbd.ttf','Arial-BoldMT',700,false],['BIZ-UDGothicR.ttc','BIZ-UDGothic',400,true],['BIZ-UDGothicB.ttc','BIZ-UDGothic-Bold',700,true]];
        const fixtures=fixtureDefinitions.filter(([file])=>fs.existsSync(path.join('C:/Windows/Fonts',file))).map(([file,postscriptName,weight,japanese],i)=>{
          const extracted=require('../desktop/font-binary.cjs').extractFont(fs.readFileSync(path.join('C:/Windows/Fonts',file)),postscriptName);
          return {id:postscriptName,family:'EngineFont'+i,weight,japanese,fsType:extracted.fsType,base64:extracted.bytes.toString('base64')};
        });
        const variableFile=path.join(root,'test-output/font-probe-v2/NotoSansJP.ttf');
        if(fs.existsSync(variableFile)){for(const weight of [400,700])fixtures.push({id:'NotoSansJP-'+weight,family:'EngineNotoJP',weight,japanese:true,variable:true,fsType:0,base64:fs.readFileSync(variableFile).toString('base64')});}
        const outlines=await window.webContents.executeJavaScript('window.runOutlineChecks('+JSON.stringify(fixtures)+')');
        if(!outlines.passed){const debug=await window.webContents.executeJavaScript('window.outlineDebug');if(debug){for(const key of ['beforeURL','afterURL']){fs.writeFileSync(path.join(evidence,'outline-'+key+'.png'),Buffer.from(debug[key].split(',')[1],'base64'));delete debug[key];}fs.writeFileSync(path.join(evidence,'outline-debug.json'),JSON.stringify(debug,null,2));}}
        report.results.push(...outlines.results);report.passed=report.results.every(item=>item.passed);
        report.performance = await window.webContents.executeJavaScript('window.runPerformanceChecks()');
        if(process.argv.includes('--large')) {
          report.largeImages = await window.webContents.executeJavaScript('window.runLargeImageChecks()');
          for(const row of report.largeImages.workerChecks) report.results.push({name:`${row.pixels/1000000}MP filtered stroke event-loop gap under500ms and one undo`,passed:row.filteredStrokePassed&&row.filteredStroke.maxEventLoopGapMs<500});
          const history=report.largeImages.history32MP;
          report.results.push({name:'32MP history trims within budget and undo restores pixels',passed:history.undoChangedPixels&&history.states.at(-1).estimatedHistoryMiB<=96&&history.states.at(-1).historyStates<=history.states.at(-2).historyStates});
          report.passed=report.results.every((item)=>item.passed);
        }
      }
      fs.writeFileSync(path.join(evidence, 'engine-check-results.json'), JSON.stringify(report, null, 2));
      console.log(JSON.stringify(report, null, 2));
      app.exit(report.passed ? 0 : 1);
    } catch (error) { console.error(error); app.exit(1); }
  });
}
