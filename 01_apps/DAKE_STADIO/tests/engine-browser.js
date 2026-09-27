import { StudioEngine, validateDocument, validateSVG } from '../src/engine.js';
import { ActiveSelection, Point } from 'fabric';

window.runEngineChecks = async () => {
  const results = [];
  const assert = (value, name, detail = '') => {
    if (!value) throw new Error(`${name}: ${detail}`);
    results.push({ name, passed: true });
  };
  const pixels = async (url) => {
    const image = new Image(); image.src = url; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
    const context = canvas.getContext('2d', { willReadFrequently: true }); context.drawImage(image, 0, 0);
    return { width: image.width, height: image.height, at: (x, y) => Array.from(context.getImageData(x, y, 1, 1).data) };
  };
  const element = document.getElementById('editor');
  const engine = new StudioEngine(element);
  window.engine = engine;
  const start = performance.now();
  try {
    engine.newDocument({ width: 320, height: 240, background: 'transparent', name: 'engine-check' });
    engine.fit(900, 680);
    assert(!engine.canUndo && !engine.canRedo, 'new document has clean history');
    const rect = engine.addShape('rect', { left: 20, top: 20, width: 80, height: 60, fill: '#ff0000' });
    const rectId = rect.id;
    assert(engine.layers.length === 1 && engine.canUndo, 'shape and history created');
    await engine.undo(); assert(engine.layers.length === 0, 'shape undo');
    await engine.redo(); assert(engine.layers.length === 1 && engine.layers[0].id === rectId, 'shape redo preserves ID');
    engine.select(rectId);
    engine.updateSelected({ angle: 12, opacity: 0.75 });
    const ellipse = engine.addShape('ellipse', { left: 130, top: 30, fill: '#00ff00' });
    engine.canvas.setActiveObject(new ActiveSelection(engine.layers, { canvas: engine.canvas }));
    const grouped = engine.group();
    assert(grouped.getObjects().length === 2 && engine.layers.length === 1, 'group two layers');
    engine.ungroup(); assert(engine.layers.length === 2, 'ungroup preserves layers');
    engine.canvas.discardActiveObject();
    const beforeGroup = await engine.exportRaster('png');
    engine.canvas.setActiveObject(new ActiveSelection(engine.layers, { canvas: engine.canvas }));
    engine.group(); engine.ungroup(); engine.canvas.discardActiveObject();
    const afterGroup = await engine.exportRaster('png');
    assert(beforeGroup === afterGroup, 'group and ungroup preserve rendered pixels');
    engine.select(rectId); await engine.duplicate(); assert(engine.layers.length === 3 && engine.selected.id !== rectId, 'duplicate creates independent ID');
    engine.deleteSelected(); assert(engine.layers.length === 2, 'delete layer');
    engine.select(rectId); engine.toggleLock(rectId);
    assert(engine.layers.find((o) => o.id === rectId).locked, 'layer locks');
    const saved = engine.serialize();
    await engine.loadDocument(saved);
    const loaded = engine.layers.find((o) => o.id === rectId);
    assert(loaded.locked && !loaded.selectable && loaded.angle === 12, 'save reload preserves lock and transform');
    const stable = JSON.stringify(engine.serialize());
    let rejected = false;
    try { await engine.loadDocument({ ...saved, width: 9000 }); } catch { rejected = true; }
    assert(rejected && JSON.stringify(engine.serialize()) === stable, 'invalid dimensions do not mutate current work');
    rejected = false;
    const corrupted = structuredClone(saved); corrupted.canvas.objects.push({ type: 'Image', src: 'data:image/png;base64,broken', width: 20, height: 20 });
    try { await engine.loadDocument(corrupted); } catch { rejected = true; }
    assert(rejected && JSON.stringify(engine.serialize()) === stable, 'failed image decode leaves current work intact', JSON.stringify({ rejected, before: JSON.parse(stable), after: engine.serialize() }));
    rejected = false;
    try { validateDocument({ ...saved, canvas: { objects: [{ type: 'Image', src: 'https://invalid.example/a.png' }] } }); } catch { rejected = true; }
    assert(rejected, 'remote project resources rejected');
    rejected = false;
    try { validateSVG('<svg xmlns="http://www.w3.org/2000/svg"><image href="https://invalid.example/a.png"/></svg>'); } catch { rejected = true; }
    assert(rejected, 'remote SVG resources rejected');
    for (const tag of ['filter', 'mask', 'pattern', 'marker', 'textPath']) {
      rejected = false;
      try { validateSVG(`<svg xmlns="http://www.w3.org/2000/svg"><${tag}/></svg>`); } catch (error) { rejected = error.message === 'UNSUPPORTED_SVG'; }
      assert(rejected, `unsupported SVG ${tag} rejects before silent loss`);
    }
    const imported = await engine.importSVG('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><path d="M10 10 C30 0 70 100 90 90 Z" fill="#3344ff"/></svg>', 'test-svg');
    assert(imported.type === 'path', 'SVG imports editable path');
    engine.setTool('node');
    const firstControl = Object.values(imported.controls)[0];
    const anchorBefore = new Point(imported.path[1][5], imported.path[1][6]).subtract(imported.pathOffset).transform(imported.calcTransformMatrix());
    const firstBefore = new Point(imported.path[0][1], imported.path[0][2]).subtract(imported.pathOffset).transform(imported.calcTransformMatrix());
    firstControl.actionHandler({}, { target: imported }, firstBefore.x + 12, firstBefore.y + 9);
    engine.commit('node-test');
    const anchorAfter = new Point(imported.path[1][5], imported.path[1][6]).subtract(imported.pathOffset).transform(imported.calcTransformMatrix());
    assert(Number.isFinite(imported.left) && Math.abs(anchorBefore.x - anchorAfter.x) < 0.001 && Math.abs(anchorBefore.y - anchorAfter.y) < 0.001, 'closed path first-node edit keeps other anchor fixed');
    engine.setTool('pen');
    const event = (x, y) => ({ scenePoint: { x, y }, e: { button: 0, clientX: x, clientY: y } });
    engine._pointerDown(event(10, 200)); engine._pointerMove(event(30, 180)); engine._pointerUp();
    engine._pointerDown(event(110, 200)); engine._pointerMove(event(130, 210)); engine._pointerUp();
    const path = engine.finishPath(false);
    assert(path.path[1][0] === 'C' && path.path[1][1] === 30 && path.path[1][2] === 180, 'pen drag creates cubic handles');
    const imageCanvas = document.createElement('canvas'); imageCanvas.width = 64; imageCanvas.height = 64;
    const imageContext = imageCanvas.getContext('2d'); imageContext.fillStyle = '#e83242'; imageContext.fillRect(0, 0, 64, 64);
    let image = await engine.importImage(imageCanvas.toDataURL(), 'pixels');
    const imageId = image.id;
    engine.setStyle({ fill: '#00ff00', brushSize: 12 }); engine.setTool('brush');
    const beforePaint = engine._history.length;
    engine._pointerDown(event(160, 120)); engine._pointerMove(event(165, 120)); engine._pointerUp();
    const painted = await pixels(image.getSrc());
    assert(painted.at(32, 32)[1] === 255 && image.dakeRaster, 'brush changes actual embedded image pixels');
    assert(engine._history.length === beforePaint + 1, 'raster gesture is one undo step');
    engine.setTool('eraser'); engine._pointerDown(event(160, 120)); engine._pointerUp();
    const erased = await pixels(image.getSrc()); assert(erased.at(32, 32)[3] === 0, 'eraser changes source alpha');
    await engine.undo(); engine.select(imageId);
    const restored = await pixels(engine.selected.getSrc()); assert(restored.at(32, 32)[3] === 255 && restored.at(32, 32)[1] === 255, 'undo restores erased raster pixels');
    await engine.applyImageAdjustment({ brightness: 0.12, saturation: -0.2, grayscale: true });
    image = engine.selected;
    const beforeFilteredStroke = image.getSrc();
    const historyBeforeFilteredStroke = engine._historyIndex;
    engine.setStyle({fill:'#2244aa'}); engine.setTool('brush');
    engine._pointerDown(event(160,120)); engine._pointerMove(event(167,120));
    const filteredStrokeTask = engine._pointerUp();
    let guardedExport = false;
    try { engine.exportSVG(); } catch (error) { guardedExport = error.message === 'BUSY'; }
    assert(engine.busy && guardedExport, 'filtered stroke blocks overlapping export while worker runs');
    await filteredStrokeTask;
    assert(engine._historyIndex === historyBeforeFilteredStroke + 1 && image.filters.length === 3 && image.getSrc() !== beforeFilteredStroke, 'filtered stroke commits changed pixels and filters in one undo step');
    await engine.undo(); image = engine.layers.find((object)=>object.id===imageId); engine.select(imageId);
    assert(image.getSrc() === beforeFilteredStroke && image.filters.length === 3, 'filtered stroke undo restores original pixels and filter chain');
    await engine.redo(); image = engine.layers.find((object)=>object.id===imageId); engine.select(imageId);
    const beforeFailedStroke = JSON.stringify(engine.serialize());
    const beforeFailedPixels = image.getSrc(true); const beforeFailedIndex = engine._historyIndex;
    const workerMethod = engine._filterInWorker; const oldStatus = engine.onStatus; let failedStatus;
    engine._filterInWorker = async (bitmap) => { bitmap.close(); throw new Error('TEST_WORKER_FAILURE'); };
    engine.onStatus = (key) => { failedStatus = key; };
    engine.setTool('brush'); engine._pointerDown(event(170,130)); await engine._pointerUp();
    engine._filterInWorker = workerMethod; engine.onStatus = oldStatus;
    assert(!engine.busy && failedStatus === 'rasterFailed' && engine._historyIndex === beforeFailedIndex && JSON.stringify(engine.serialize()) === beforeFailedStroke && image.getSrc(true) === beforeFailedPixels, 'failed filtered stroke rolls back source, visible pixels, filters and history');
    const workerPixels = image.getSrc(true);
    image.applyFilters();
    assert(workerPixels === image.getSrc(true), 'worker chain output exactly matches Fabric CPU filter pixels');
    const filteredPreview = await engine.exportRaster('png');
    const filtered = engine.serialize(); await engine.loadDocument(filtered); engine.select(imageId);
    assert(filteredPreview === await engine.exportRaster('png'), 'worker-filtered project reload preserves exact rendered pixels');
    assert(engine.selected.filters.length === 3 && engine.selected.adjustments.brightness === 0.12, 'editable adjustments survive project reload');
    const sourceBeforeFilterReset = engine.selected.getSrc(); await engine.applyImageAdjustment({});
    assert(engine.selected.getSrc() === sourceBeforeFilterReset && !engine.selected.filters.length, 'adjustment reset preserves original source pixels');
    engine.setZoom(0.5); const pngA = await engine.exportRaster('png');
    engine.setZoom(2.5); const pngB = await engine.exportRaster('png');
    assert(pngA === pngB, 'PNG export invariant under zoom');
    const png = await pixels(pngA); assert(png.width === 320 && png.height === 240, 'PNG export exact document dimensions');
    const jpg = await pixels(await engine.exportRaster('jpeg')); assert(jpg.at(319, 239)[3] === 255, 'JPEG export has opaque background');
    assert((await engine.exportRaster('webp')).startsWith('data:image/webp;base64,'), 'WebP export encodes WebP');
    engine.setZoom(0.2); const svgA = engine.exportSVG(); engine.setZoom(3); const svgB = engine.exportSVG();
    const normalizeSVG = (svg) => svg.replace(/CLIPPATH_\d+/g, 'CLIPPATH_ID');
    assert(normalizeSVG(svgA) === normalizeSVG(svgB) && !svgA.includes('<!DOCTYPE') && svgA.includes('<path') && svgA.includes('data:image/png;base64,'), 'SVG export retains vector and embedded pixels without viewport or DOCTYPE', JSON.stringify({a:svgA.slice(0,1200),b:svgB.slice(0,1200)}));
    engine.select(imageId); await engine.rasterizeSelected();
    assert(engine.selected.type === 'image', 'rasterize produces editable image layer');
    engine.cropDocument({ left: 10, top: 20, width: 180, height: 140 });
    assert(engine.width === 180 && engine.height === 140, 'crop document bounds');
    await engine.undo(); assert(engine.width === 320 && engine.height === 240, 'crop undo restores dimensions');
    const finalDocument = engine.serialize();
    const screenshot = await engine.exportRaster('png');
    return { passed: true, results, durationMs: Math.round(performance.now() - start), finalDocument, screenshot, svg: engine.exportSVG() };
  } catch (error) {
    return { passed: false, results, error: error.stack, durationMs: Math.round(performance.now() - start) };
  }
};

window.prepareNativeInputChecks = async () => {
  const engine = window.engine;
  engine.newDocument({ width: 320, height: 240, name: 'native-input', background: '#ffffff' });
  engine.fit(900, 680);
  const source = document.createElement('canvas'); source.width = 64; source.height = 64;
  const context = source.getContext('2d'); context.fillStyle = '#ff0000'; context.fillRect(0, 0, 64, 64);
  const image = await engine.importImage(source.toDataURL(), 'native-paint');
  engine.setStyle({ fill: '#0000ff', brushSize: 12 }); engine.setTool('brush'); engine.canvas.renderAll();
  const box = engine.canvas.upperCanvasEl.getBoundingClientRect();
  const viewport = engine.canvas.viewportTransform;
  const screenPoint = (x, y) => ({ x: Math.round(box.left + x * viewport[0] + viewport[4]), y: Math.round(box.top + y * viewport[3] + viewport[5]) });
  window.nativeImageId = image.id;
  return { start: screenPoint(160, 120), end: screenPoint(170, 120), pen1: screenPoint(20, 200), pen1handle: screenPoint(40, 180), pen2: screenPoint(110, 200), pen2handle: screenPoint(130, 210) };
};

window.checkNativePixels = () => {
  const image = window.engine.layers.find((object) => object.id === window.nativeImageId);
  const canvas = document.createElement('canvas'); canvas.width = 64; canvas.height = 64;
  const context = canvas.getContext('2d'); context.drawImage(image._originalElement, 0, 0);
  return Array.from(context.getImageData(32, 32, 1, 1).data);
};

window.runPerformanceChecks = async () => {
  const engine = window.engine;
  const start = performance.now();
  engine.newDocument({ width: 2048, height: 1536, name: '3mp-response', background: 'transparent' });
  engine.setTool('brush');
  const frameDelays = [];
  let running = true; let previous = performance.now();
  const measure = () => { const now = performance.now(); frameDelays.push(now - previous); previous = now; if (running) requestAnimationFrame(measure); };
  requestAnimationFrame(measure);
  const gestureTimes = [];
  for (let gesture = 0; gesture < 6; gesture++) {
    const t = performance.now();
    const x = 100 + gesture * 200; const y = 200 + gesture * 100;
    engine._pointerDown({ scenePoint: {x,y}, e: {button:0} });
    for (let step = 0; step < 20; step++) engine._pointerMove({ scenePoint: {x:x+step*7,y:y+step*4}, e:{button:0} });
    engine._pointerUp();
    gestureTimes.push(Math.round(performance.now() - t));
    await new Promise((resolve) => requestAnimationFrame(resolve));
  }
  const undoStart = performance.now(); await engine.undo(); const undoMs = Math.round(performance.now() - undoStart);
  const exportStart = performance.now(); await engine.exportRaster('png'); const exportMs = Math.round(performance.now() - exportStart);
  running = false;
  return { document: '2048x1536', rasterGestures: 6, gestureTimesMs: gestureTimes, undoMs, exportMs, maxFrameGapMs: Math.round(Math.max(...frameDelays)), totalMs: Math.round(performance.now()-start) };
};

window.runLargeImageChecks = async () => {
  const engine = window.engine;
  const result = { workerChecks: [], history32MP: null };
  const heartbeat = async (action) => {
    let previous = performance.now(); let count = 0; let maxGap = 0;
    const timer = setInterval(() => { const now = performance.now(); maxGap = Math.max(maxGap, now - previous); previous = now; count++; }, 10);
    const start = performance.now(); await action();
    maxGap = Math.max(maxGap, performance.now()-previous); clearInterval(timer);
    return { elapsedMs: Math.round(performance.now()-start), eventLoopTicks: count, maxEventLoopGapMs: Math.round(maxGap) };
  };
  for (const [width,height] of [[4000,3000],[8000,4000]]) {
    engine.newDocument({width,height,name:'worker-size-check',background:'transparent'});
    const noise = document.createElement('canvas'); noise.width = width/8; noise.height = height/8;
    const nc = noise.getContext('2d'); const nd = nc.createImageData(noise.width,noise.height);
    let seed = 0x5139487;
    for (let i=0;i<nd.data.length;i+=4) { seed ^= seed<<13; seed ^= seed>>>17; seed ^= seed<<5; nd.data[i]=seed&255;nd.data[i+1]=(seed>>>8)&255;nd.data[i+2]=(seed>>>16)&255;nd.data[i+3]=255; }
    nc.putImageData(nd,0,0);
    const canvas = document.createElement('canvas'); canvas.width=width;canvas.height=height;
    const ctx=canvas.getContext('2d');ctx.imageSmoothingEnabled=false;ctx.drawImage(noise,0,0,width,height);
    const imported=await engine.importImage(canvas.toDataURL(),'large-pixels');
    const source=imported.getSrc();
    const worker=await heartbeat(()=>engine.applyImageAdjustment({brightness:0.08,contrast:0.14,saturation:-0.2,grayscale:true}));
    const busyGuard = await (async()=>{
      const job=engine.applyImageAdjustment({brightness:0.09,contrast:0.12,saturation:-0.3});
      let rejected=false;try{engine.newDocument({width:10,height:10,name:'blocked',background:'#fff'});}catch(error){rejected=error.message==='BUSY';}
      await job;return rejected&&engine.width===width;
    })();
    const cpuStart=performance.now();imported.applyFilters();const cpuMs=Math.round(performance.now()-cpuStart);
    const serialized=engine.serialize();
    const reload=await heartbeat(()=>engine.loadDocument(serialized));
    engine.select(imported.id);
    const sourcePreserved=engine.selected.getSrc()===source;
    const beforeStrokeIndex=engine._historyIndex;
    engine.setStyle({fill:'#ff22bb',brushSize:60});engine.setTool('brush');
    const filteredStroke=await heartbeat(()=>{
      engine._pointerDown({scenePoint:{x:1000,y:900},e:{button:0}});
      engine._pointerMove({scenePoint:{x:1050,y:930},e:{button:0}});
      return engine._pointerUp();
    });
    const filteredStrokePassed=engine._historyIndex===beforeStrokeIndex+1&&engine.selected.filters.length===3&&engine.selected.getSrc()!==source;
    result.workerChecks.push({width,height,pixels:width*height,worker,referenceCpuMs:cpuMs,reload,busyGuardPassed:busyGuard,sourcePreserved,filteredStroke,filteredStrokePassed});
    await engine.applyImageAdjustment({});
    if(width===8000){
      const states=[];engine.setStyle({fill:'#ff00aa',brushSize:60});engine.setTool('brush');
      for(let i=0;i<22;i++){
        const started=performance.now();const x=700+i*100,y=600+i*45;
        engine._pointerDown({scenePoint:{x,y},e:{button:0}});engine._pointerMove({scenePoint:{x:x+45,y:y+20},e:{button:0}});engine._pointerUp();
        const pooledBytes=Array.from(engine._images.values()).reduce((total,item)=>total+item.length*2,0)+engine._history.reduce((total,item)=>total+item.length*2,0);
        states.push({gesture:i+1,historyStates:engine._history.length,sourceVersions:engine._images.size,estimatedHistoryMiB:Math.round(pooledBytes/1048576*10)/10,gestureMs:Math.round(performance.now()-started)});
        await new Promise(resolve=>setTimeout(resolve,0));
        if(i>1&&states.at(-1).historyStates<=states.at(-2).historyStates)break;
      }
      const lastSource=engine.selected.getSrc();await engine.undo();const undoSource=engine.layers.find(o=>o.id===imported.id).getSrc();
      result.history32MP={sourceDataURLMiB:Math.round(source.length/1048576*10)/10,states,undoChangedPixels:lastSource!==undoSource,budgetMiB:96,minimumStates:2};
    }
  }
  return result;
};

window.runFoundationChecks=async()=>{
  const e=window.engine,results=[];
  const check=(value,name,detail)=>{results.push({name,passed:!!value,...(detail?{detail}:{})});if(!value)throw new Error(name+': '+JSON.stringify(detail));};
  const pixels=async url=>{const image=new Image();image.src=url;await image.decode();const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const ctx=c.getContext('2d');ctx.drawImage(image,0,0);return{at:(x,y)=>Array.from(ctx.getImageData(x,y,1,1).data),data:ctx.getImageData(0,0,c.width,c.height).data};};
  const reset=()=>{e.newDocument({width:300,height:200,background:'transparent',name:'foundation'});e.fit(900,680);};
  try{
    reset();e.updateDocumentSettings({dpi:300,unit:'mm',bleed:5,safe:8,snap:{grid:true,gridSize:10}});
    e.addGuide('x',80);e.addGuide('y',70);
    const meta=structuredClone(e.meta),data=e.serialize();
    check(data.version===2&&data.meta.dpi===300&&data.meta.guides.vertical[0]===80,'v2 contains DPI units bleed safe and guides');
    await e.loadDocument(data);check(JSON.stringify(meta)===JSON.stringify(e.meta),'v2 metadata reopens exactly');
    const v1={...data,version:1};delete v1.meta;delete v1.fonts;await e.loadDocument(v1);
    check(e.meta.dpi===96&&e.meta.unit==='px','v1 projects load with compatible production defaults');
    const r=e.addShape('rect',{left:10,top:20,width:40,height:30,fill:'#f00'});
    const r2=e.addShape('rect',{left:85,top:60,width:40,height:30,fill:'#0f0'});
    const r3=e.addShape('rect',{left:230,top:80,width:40,height:30,fill:'#00f'});
    e.selectMany([r.id,r2.id,r3.id]);e.alignSelected('top');e.canvas.discardActiveObject();
    check(e.layers.every(o=>Math.abs(o.top-r.top)<.001),'multiple layers align by actual bounds');
    e.selectMany([r.id,r2.id,r3.id]);e.distributeSelected('horizontal');e.canvas.discardActiveObject();
    check(Math.abs((r2.left-r.left)-(r3.left-r2.left))<.01,'three layers distribute equal gaps');
    e.renameLayer(r.id,'renamed');e.moveLayer(r.id,2);
    check(e.layers.at(-1).id===r.id&&r.name==='renamed','layer rename and indexed reorder');
    e.select(r.id);e.select(r2.id,{additive:true});
    check(e.canvas.getActiveObjects().length===2,'additive layer selection retains earlier selection');
    e.canvas.discardActiveObject();e.updateDocumentSettings({snap:{enabled:true,artboard:false,objects:false,guides:true,grid:false}});e.addGuide('x',50);
    r.set({left:49,top:30});r.setCoords();e._snapMoving(r,{});
    check(Math.abs(r.getBoundingRect().left-50)<.01&&e.view.snapLines.some(l=>l.axis==='x'),'guide snapping moves bounds and provides overlay lines');

    const booleanPixels=async(operation)=>{
      reset();const a=e.addShape('rect',{left:20,top:20,width:100,height:100,fill:'#f00'}),b=e.addShape('rect',{left:70,top:20,width:100,height:100,fill:'#00f'});
      e.selectMany([a.id,b.id]);const history=e._historyIndex;await e.booleanSelected(operation);
      check(e._historyIndex===history+1&&e.layers.every(o=>o.type==='path'),'boolean '+operation+' commits editable paths in one undo');
      const result=await pixels(await e.exportRaster());const alphas=[result.at(40,50)[3],result.at(90,50)[3],result.at(150,50)[3]];
      const expected={union:[255,255,255],subtract:[255,0,0],intersect:[0,255,0],divide:[255,255,0],compound:[255,0,255]}[operation];
      check(alphas.every((a,i)=>a===expected[i]),'boolean '+operation+' has correct inside/outside geometry',{alphas,expected});
      const png=await e.exportRaster(),saved=e.serialize();await e.loadDocument(saved);
      check(await e.exportRaster()===png,'boolean '+operation+' keeps pixels through reopen');
    };
    for(const operation of ['union','subtract','intersect','divide','compound'])await booleanPixels(operation);
    e.select(e.layers[0].id);await e.splitCompound();check(e.layers.length===2,'compound splits into independently editable contours');
    reset();const circle=e.addShape('ellipse',{left:50,top:30,rx:65,ry:40,fill:'#f00'}),box=e.addShape('rect',{left:90,top:0,width:30,height:160,fill:'#00f'});
    e.selectMany([circle.id,box.id]);await e.booleanSelected('subtract');
    check(e.layers[0].path.some(c=>c[0]==='C'),'curved boolean retains cubic Bezier segments');
    reset();e.addShape('rect',{left:10,top:10,width:20,height:20});e.addShape('rect',{left:100,top:100,width:20,height:20});
    e.selectMany(e.layers.map(o=>o.id));const beforeEmpty=JSON.stringify(e.serialize());let rejected=false;
    try{await e.booleanSelected('intersect');}catch(error){rejected=error.message==='VECTOR_EMPTY';}
    check(rejected&&JSON.stringify(e.serialize())===beforeEmpty,'empty boolean leaves all original objects intact');

    reset();
    const source=document.createElement('canvas');source.width=200;source.height=150;
    const sourceCtx=source.getContext('2d');sourceCtx.fillStyle='#ff8800';sourceCtx.fillRect(0,0,200,150);sourceCtx.fillStyle='#0088ff';sourceCtx.fillRect(100,0,100,150);
    const image=await e.importImage(source.toDataURL(),'original photo');e.updateSelected({left:20,top:20,scaleX:1,scaleY:1});
    const mask=e.addShape('ellipse',{left:60,top:40,rx:50,ry:50,fill:'#000',strokeWidth:0});
    const imageId=image.id,maskId=mask.id,originalSrc=image.getSrc(),beforeMask=JSON.stringify(e.serialize());
    e.selectMany([image.id,mask.id]);const group=await e.createMask();
    let px=await pixels(await e.exportRaster());check(px.at(100,80)[3]===255&&px.at(25,25)[3]===0,'non-destructive shape mask clips pixels without deleting photo');
    check(group.getObjects()[0].getSrc()===originalSrc,'mask retains original embedded image bytes');
    const maskedPNG=await e.exportRaster(),savedMask=e.serialize();
    check(/clipPath/.test(e.exportSVG()),'SVG contains editable vector clipping geometry');
    await e.loadDocument(savedMask);check(await e.exportRaster()===maskedPNG,'mask project reload preserves exact PNG pixels');
    e.select(e.layers[0].id);await e.toggleMask();px=await pixels(await e.exportRaster());check(px.at(25,25)[3]===255,'mask can be disabled and reveals original image');
    const disabled=e.serialize();await e.loadDocument(disabled);e.select(e.layers[0].id);await e.toggleMask();
    check(await e.exportRaster()===maskedPNG,'disabled mask reopens and can be enabled without lost clip');
    await e.releaseMask();e.canvas.discardActiveObject();
    check(e.layers.length===2&&e.layers.some(o=>o.id===imageId&&o.getSrc()===originalSrc)&&e.layers.some(o=>o.id===maskId),'release restores original image and shape identifiers');
    const released=JSON.stringify(e.serialize());await e.undo();check(e.layers.length===1&&e.layers[0].dakeMask.enabled,'mask release is undoable');
    await e.redo();check(JSON.stringify(e.serialize())===released,'mask release redo retains geometry');
    await e.loadDocument(savedMask);e.select(e.layers[0].id);await e.rasterizeSelected();
    px=await pixels(await e.exportRaster());check(px.at(100,80)[3]===255&&px.at(25,25)[3]===0&&e.layers[0].type==='image','rasterization respects mask while retaining pixel editing');
    await e.loadDocument(savedMask);e.select(e.layers[0].id);e.updateSelected({left:100,top:60,angle:15,scaleX:1.1,scaleY:.8});
    const transformed=e.selected.calcTransformMatrix().slice();await e.releaseMask();e.canvas.discardActiveObject();
    check(e.layers.every(o=>Number.isFinite(o.left)&&Number.isFinite(o.angle))&&Math.abs(e.layers[0].angle-15)<.001,'mask release carries current group transform into source objects');

    reset();await e.importSVG('<svg xmlns="http://www.w3.org/2000/svg" width="300" height="200"><path d="M20 60 C50 10 120 10 150 60 L150 130 L20 130 Z" fill="#f00"/></svg>','curve');
    const path=e.selected;const beforeNode=await pixels(await e.exportRaster());e.setTool('node');e.activeNodeIndex=0;e.editNode('insert');
    const afterNode=await pixels(await e.exportRaster());
    let delta=0;for(let i=0;i<beforeNode.data.length;i++)delta+=Math.abs(beforeNode.data[i]-afterNode.data[i]);
    check(path.path.filter(c=>c[0]==='C').length===2&&delta<3000,'node insertion splits cubic curve without changing shape',{delta});
    e.editNode('smooth',1);check(path.path[1][0]==='C'&&path.path[2][0]==='C','smooth node creates continuous cubic handles');
    e.editNode('corner',1);check(path.path[1][3]===path.path[1][5]&&path.path[2][1]===path.path[1][5],'corner removes adjacent handles');
    e.editNode('remove',1);check(path.path.length===5,'node removal changes path topology');
    e.editNode('open');check(!path.path.some(c=>c[0]==='Z'),'closed contour can be opened');e.editNode('close');check(path.path.at(-1)[0]==='Z','open contour can be closed');
    e.editNode('smooth',0);const seamCount=path.path.length;e.editNode('smooth',0);
    check(path.path.length===seamCount&&path.path.at(-2).at(-2)===path.path[0][1]&&path.path.at(-2).at(-1)===path.path[0][2],'repeated first-node smoothing retains one joined closing seam');
    e.select(null);check(e.activeNodeIndex===null,'node selection resets when leaving the path');e.select(path.id);
    const nodeDoc=e.serialize();await e.loadDocument(nodeDoc);check(JSON.stringify(e.layers[0].path)===JSON.stringify(path.path),'node topology survives save and reload');
    reset();const im=await e.importImage(source.toDataURL(),'dpi');e.updateDocumentSettings({dpi:300});e.updateSelected({scaleX:2,scaleY:1.5});
    const dpi=e.getEffectiveDPI(im);check(Math.abs(dpi.x-150)<.001&&Math.abs(dpi.y-200)<.001,'effective image DPI reflects actual transformed pixel density');

    e.updateDocumentSettings({bleed:20,safe:30,guides:{vertical:[10,120,290],horizontal:[20,170]}});
    e.cropDocument({left:100,top:80,width:20,height:15});
    validateDocument(e.serialize());check(e.meta.guides.vertical.every(x=>x<=20)&&2*(e.meta.bleed+e.meta.safe)<15,'small crop keeps production metadata inside new document bounds');
    await e.loadDocument(savedMask);e.select(e.layers[0].id);await e.duplicate();
    const copy=e.selected,originalGroup=e.layers[0];check(copy.dakeMask.source.id!==originalGroup.dakeMask.source.id,'duplicated mask receives independent source shape identifier');
    await e.ungroup();check(e.layers.length===3&&e.layers.filter(o=>o.type==='image').length===1,'ungroup mask restores original clipping shape');
    reset();const locked=e.addShape('rect');e.toggleLock(locked.id);const free=e.addShape('rect');e.selectMany([locked.id,free.id]);check(e.canvas.getActiveObjects().length===1,'layer multiple selection respects locked objects');
    reset();const first=e.addShape('rect',{left:20,top:20,width:80,height:80}),second=e.addShape('ellipse',{left:40,top:20,rx:40,ry:40});e.selectMany([first.id,second.id]);
    const operation=e.booleanSelected('union');let undoBusy=false;try{await e.undo();}catch(error){undoBusy=error.message==='BUSY';}await operation;check(undoBusy,'vector worker rejects overlapping undo');
    return {passed:true,results};
  }catch(error){results.push({name:error.message,passed:false,stack:error.stack});return{passed:false,results};}
};


window.runOutlineChecks=async fixtures=>{
  const e=window.engine,results=[];
  const check=(value,name,detail)=>{results.push({name,passed:!!value,detail});if(!value)throw new Error(name+': '+JSON.stringify(detail));};
  const pixels=async url=>{const image=new Image();image.src=url;await image.decode();const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const ctx=c.getContext('2d');ctx.drawImage(image,0,0);return ctx.getImageData(0,0,c.width,c.height).data;};
  const compare=(a,b)=>{let overlap=0,union=0,areaA=0,areaB=0;for(let i=3;i<a.length;i+=4){const aa=a[i]>64,bb=b[i]>64;if(aa)areaA++;if(bb)areaB++;if(aa&&bb)overlap++;if(aa||bb)union++;}return{iou:overlap/union,areaA,areaB,areaRatio:areaB/areaA};};
  try{
    for(const fixture of fixtures){
      const buffer=Uint8Array.from(atob(fixture.base64),c=>c.charCodeAt(0)).buffer;
      const face=new FontFace(fixture.family,buffer,{weight:fixture.variable?'100 900':String(fixture.weight),style:'normal'});await face.load();document.fonts.add(face);
      for(const tracking of [0,80]){
        e.newDocument({width:600,height:320,background:'transparent',name:'outline'});e.fit(900,680);
        const object=e.addText(fixture.japanese?'日本語の制作\nDAKE STADIO':'AVATAR office\nDAKE STADIO');
        e.updateSelected({left:35,top:25,fontFamily:fixture.family,fontSize:44,fontWeight:fixture.weight,fill:'#123456',charSpacing:tracking,lineHeight:1.35,textAlign:tracking?'center':'left',underline:!!tracking});
        const beforeURL=await e.exportRaster();const before=await pixels(beforeURL),history=e._historyIndex,text=object.text;
        let lastTick=performance.now(),maxGap=0;const heartbeat=setInterval(()=>{const now=performance.now();maxGap=Math.max(maxGap,now-lastTick);lastTick=now;},16),started=performance.now();
        const path=await e.outlineText(buffer,{family:fixture.family,source:'local',id:fixture.id,fsType:fixture.fsType});
        const duration=performance.now()-started;await new Promise(resolve=>setTimeout(resolve,20));clearInterval(heartbeat);
        check(maxGap<500,'font worker keeps foreground responsive '+fixture.id+' '+tracking,{durationMs:Math.round(duration),maxEventLoopGapMs:Math.round(maxGap)});
        const afterURL=await e.exportRaster();const after=await pixels(afterURL),metric=compare(before,after);window.outlineDebug={beforeURL,afterURL,object:object.toObject(),path:path.toObject()};
        check(path.type==='path'&&path.path.some(command=>command[0]==='C'||command[0]==='Q')&&e._historyIndex===history+1,'font '+fixture.id+' tracking '+tracking+' outlines into editable curves in one step');
        check(metric.iou>.82&&metric.areaRatio>.9&&metric.areaRatio<1.12,'font '+fixture.id+' tracking '+tracking+' retains glyph layout and weight',metric);
        const svg=e.exportSVG();check(!/<text[\s>]/.test(svg),'outline SVG has no font dependency '+fixture.id+' '+tracking);
        const png=await e.exportRaster(),saved=e.serialize();await e.loadDocument(saved);
        check(await e.exportRaster()===png,'outline project reload preserves exact pixels '+fixture.id+' '+tracking);
        e.select(e.layers[0].id);
      }
      e.newDocument({width:600,height:200,background:'transparent',name:'outline-rollback'});e.addText('test');e.updateSelected({fontFamily:fixture.family,fontWeight:fixture.weight});
      const stable=JSON.stringify(e.serialize());let rejected=false;
      try{await e.outlineText(new ArrayBuffer(5),{family:fixture.family});}catch{rejected=true;}
      check(rejected&&JSON.stringify(e.serialize())===stable,'invalid font preserves editable text '+fixture.id);
      document.fonts.delete(face);
    }
    return{passed:true,results};
  }catch(error){results.push({name:error.message,passed:false,stack:error.stack});return{passed:false,results};}
};

