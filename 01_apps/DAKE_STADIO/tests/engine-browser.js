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
