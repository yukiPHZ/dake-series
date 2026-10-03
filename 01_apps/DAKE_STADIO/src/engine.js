import {validateFontIdentity} from '../desktop/font-identity.cjs';
import {
  Canvas, StaticCanvas, FabricObject, FabricImage, Rect, Ellipse, Triangle, Line,
  Path, Textbox, Group, ActiveSelection, Point, controlsUtils, util, filters, config,
  loadSVGFromString, Canvas2dFilterBackend, setFilterBackend,
} from 'fabric';
import UI_TEXT from './ui-text.json';
import {installImaging, validateImageEditing, IMAGE_EDIT_PROPERTIES} from './engine-imaging.js';
import {installComponents, validateComponents, COMPONENT_PROPERTIES} from './engine-components.js';
import {installFoundation, normalizeMeta, validateV2Extras, FOUNDATION_PROPERTIES} from './engine-foundation.js';

// Rendering stays on this machine. The CPU backend is the serialization
// reference; supported live adjustments run in a dedicated local worker.
setFilterBackend(new Canvas2dFilterBackend());
// Retain production vector placement through repeated save/load cycles.
config.NUM_FRACTION_DIGITS = 12;
const CUSTOM = ['id', 'name', 'locked', 'dakeRaster', 'adjustments', 'objectCaching', ...FOUNDATION_PROPERTIES, ...COMPONENT_PROPERTIES, ...IMAGE_EDIT_PROPERTIES];
FabricObject.customProperties = CUSTOM;
const MAX_AREA = 32_000_000;
const MAX_SIDE = 8192;
const MAX_OBJECTS = 2000;
const MAX_HISTORY = 50;
const MAX_HISTORY_BYTES = 96 * 1024 * 1024;
const IMAGE_DATA = /^data:image\/(?:png|jpeg|webp|gif|bmp);base64,/i;
const TYPES = new Set(['rect', 'ellipse', 'triangle', 'line', 'path', 'textbox', 'i-text', 'itext', 'text', 'group', 'image', 'circle', 'polygon', 'polyline']);
const FILTER_TYPES = new Set(['Brightness', 'Contrast', 'Saturation', 'Grayscale']);
const identity = () => [1, 0, 0, 1, 0, 0];
const uid = () => globalThis.crypto?.randomUUID?.() || `layer-${Date.now()}-${Math.random().toString(36).slice(2)}`;
const number = (n, fallback = 0) => Number.isFinite(Number(n)) ? Number(n) : fallback;
const clamp = (n, min, max) => Math.min(max, Math.max(min, number(n)));
const layerName = (key) => {
  const lower = String(key || '').toLowerCase();
  const aliases = { textbox: 'text', 'i-text': 'text', itext: 'text', circle: 'ellipse', polygon: 'path', polyline: 'path' };
  return UI_TEXT.layerNames?.[aliases[lower] || lower] || UI_TEXT.layerNames?.path;
};
const rasterCanvas = (width, height) => {
  const el = document.createElement('canvas');
  el.width = width;
  el.height = height;
  return el;
};

export function validateDimensions(width, height) {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 ||
      width > MAX_SIDE || height > MAX_SIDE || width * height > MAX_AREA) {
    throw new Error('DOCUMENT_TOO_LARGE');
  }
}

// Reject remote resources before Fabric has any opportunity to resolve them.
export function validateDocument(data, {skipComponents=false} = {}) {
  if (!data || data.format !== 'dake-stadio' || ![1, 2, 3].includes(data.version) ||
      !data.canvas || !Array.isArray(data.canvas.objects) || typeof data.name !== 'string' ||
      typeof data.background !== 'string') throw new Error('INVALID_DOCUMENT');
  validateDimensions(data.width, data.height);
  validateV2Extras(data);
  if (data.name.length > 1000 || data.background.length > 128 ||
      !CSS.supports('color', data.background)) throw new Error('INVALID_DOCUMENT');
  let count = 0;
  const walk = (value, depth = 0) => {
    if (depth > 60) throw new Error('INVALID_DOCUMENT');
    if (!value || typeof value !== 'object') return;
    if (Array.isArray(value) && value.length > 100_000) throw new Error('INVALID_DOCUMENT');
    for (const [key, item] of Object.entries(value)) {
      if (key === 'dakeComponent') continue;
      if (['__proto__', 'prototype', 'constructor'].includes(key)) throw new Error('INVALID_DOCUMENT');
      if (typeof item === 'number' && (!Number.isFinite(item) || Math.abs(item) > 1e8)) throw new Error('INVALID_DOCUMENT');
      if (['src', 'source', 'href', 'xlink:href'].includes(key) && typeof item === 'string' && !IMAGE_DATA.test(item)) {
        throw new Error('INVALID_DOCUMENT');
      }
      if (['fill', 'stroke', 'background', 'backgroundColor'].includes(key) && typeof item === 'string' && /url\s*\(/i.test(item)) throw new Error('INVALID_DOCUMENT');
      walk(item, depth + 1);
    }
  };
  const objects = (items, depth = 0) => {
    if (depth > 20) throw new Error('INVALID_DOCUMENT');
    for (const object of items) {
      if (++count > MAX_OBJECTS || !object || !TYPES.has(String(object.type).toLowerCase())) throw new Error('INVALID_DOCUMENT');
      if(String(object.type).toLowerCase()==='image')validateImageEditing(object);
      if (object.excludeFromExport) throw new Error('INVALID_DOCUMENT');
      if (object.path && (!Array.isArray(object.path) || object.path.length > 40_000)) throw new Error('INVALID_DOCUMENT');
      if (object.filters && (!Array.isArray(object.filters) || object.filters.length > 4 || object.filters.some((filter) => !FILTER_TYPES.has(filter?.type)))) throw new Error('INVALID_DOCUMENT');
      if (object.resizeFilter) throw new Error('INVALID_DOCUMENT');
      if (object.objects) {
        if (!Array.isArray(object.objects)) throw new Error('INVALID_DOCUMENT');
        objects(object.objects, depth + 1);
      }
      if (object.clipPath) objects([object.clipPath], depth + 1);
      if(object.dakeMask){
        if(String(object.type).toLowerCase()!=='group'||typeof object.dakeMask.enabled!=='boolean'||!object.dakeMask.source||!['path','rect','ellipse','circle','triangle','polygon'].includes(String(object.dakeMask.source.type).toLowerCase())||(object.dakeMask.enabled&&!object.clipPath))throw new Error('INVALID_DOCUMENT');
        objects([object.dakeMask.source], depth + 1);
      }
    }
  };
  walk(data.canvas);
  objects(data.canvas.objects);
  if(!skipComponents){validateComponents(data,child=>validateDocument(child,{skipComponents:true}));validateFontIdentity(data);}
  return data;
}

export function validateSVG(source) {
  if (typeof source !== 'string' || source.length > 12 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error('UNSAFE_SVG');
  const xml = new DOMParser().parseFromString(source, 'image/svg+xml');
  if (xml.querySelector('parsererror') || xml.documentElement.localName !== 'svg') throw new Error('UNSAFE_SVG');
  const nodes = Array.from(xml.querySelectorAll('*'));
  if (nodes.length > 10_000) throw new Error('UNSAFE_SVG');
  for (const node of nodes) {
    if (['script', 'foreignobject', 'iframe', 'audio', 'video', 'animate', 'set', 'animatetransform'].includes(node.localName.toLowerCase())) throw new Error('UNSAFE_SVG');
    if (['filter', 'mask', 'pattern', 'marker', 'textpath'].includes(node.localName.toLowerCase())) throw new Error('UNSUPPORTED_SVG');
    for (const attr of Array.from(node.attributes)) {
      const value = attr.value.trim();
      if (['filter', 'mask', 'marker-start', 'marker-mid', 'marker-end'].includes(attr.name.toLowerCase()) && value !== 'none') throw new Error('UNSUPPORTED_SVG');
      if (attr.name.toLowerCase() === 'style' && /(?:filter|mask|marker-(?:start|mid|end))\s*:/i.test(value)) throw new Error('UNSUPPORTED_SVG');
      if (/^on/i.test(attr.name)) throw new Error('UNSAFE_SVG');
      if (['href', 'src', 'xlink:href'].includes(attr.name.toLowerCase()) && value && !value.startsWith('#') && !IMAGE_DATA.test(value)) throw new Error('UNSAFE_SVG');
      for (const match of value.matchAll(/url\s*\(\s*['"]?([^)'"\s]+)/gi)) if (!match[1].startsWith('#')) throw new Error('UNSAFE_SVG');
    }
    if (node.localName.toLowerCase() === 'style') {
      if (/(?:filter|mask|marker-(?:start|mid|end))\s*:/i.test(node.textContent)) throw new Error('UNSUPPORTED_SVG');
      if (/@import|@font-face|expression\s*\(/i.test(node.textContent)) throw new Error('UNSAFE_SVG');
      for (const match of node.textContent.matchAll(/url\s*\(\s*['"]?([^)'"\s]+)/gi)) if (!match[1].startsWith('#')) throw new Error('UNSAFE_SVG');
    }
  }
  return new XMLSerializer().serializeToString(xml);
}

export class StudioEngine {
  constructor(canvasElement, { onChange = () => {}, onSelection = () => {}, onStatus = () => {} } = {}) {
    this.onChange = onChange;
    this.onSelection = onSelection;
    this.onStatus = onStatus;
    this.canvas = new Canvas(canvasElement, {
      preserveObjectStacking: true, selectionColor: '#66d9c622', selectionBorderColor: '#36b9a7',
      selectionLineWidth: 1, stopContextMenu: true, fireRightClick: false,
      enableRetinaScaling: true, perPixelTargetFind: false,
    });
    this.style = { fill: '#28b6a3', stroke: '#1b3b45', strokeWidth: 0, brushSize: 28 };
    this.width = 1200;
    this.height = 800;
    this.background = '#ffffff';
    this.name = '';
    this.activeTool = 'select';
    this._busy = false;
    this._history = [];
    this._historyIndex = -1;
    this._images = new Map();
    this._imageTokens = new Map();
    this._nextImageToken = 1;
    this._controls = new WeakMap();
    this._nodes = [];
    this._preview = null;
    this._stroke = null;
    this._strokeTask = null;
    this._drag = null;
    this._pendingNode = null;
    this._revision = 0;
    this._worker = null;
    this._workerJob = null;
    this._workerSequence = 0;
    this._initFoundation();
    this.canvas.on('selection:created', () => this._selectionChanged());
    this.canvas.on('selection:updated', () => this._selectionChanged());
    this.canvas.on('selection:cleared', () => this._selectionChanged());
    this.canvas.on('object:modified', () => this.commit('transform'));
    this.canvas.on('text:changed', () => this._emit('text-preview'));
    this.canvas.on('text:editing:exited', () => this.commit('text'));
    this.canvas.on('mouse:down:before', () => {
      // Fabric clears selection when hit testing is disabled. Preserve the
      // explicitly chosen raster layer before that internal selection pass.
      this._gestureTarget = this._busy || ['brush', 'eraser'].includes(this.activeTool) ? this.selected : null;
    });
    this.canvas.on('mouse:down', (event) => this._pointerDown(event));
    this.canvas.on('mouse:move', (event) => this._pointerMove(event));
    this.canvas.on('mouse:up', () => this._pointerUp());
    this.canvas.on('mouse:dblclick', () => { if (this.activeTool === 'pen') this.finishPath(false); });
    this.canvas.on('mouse:wheel', ({ e }) => {
      e.preventDefault();
      e.stopPropagation();
      if (this._busy) return;
      if (e.ctrlKey || e.metaKey) {
        const point = new Point(e.offsetX, e.offsetY);
        this.canvas.zoomToPoint(point, clamp(this.zoom * Math.pow(0.999, e.deltaY), 0.02, 16));
      } else {
        const vpt = this.canvas.viewportTransform.slice();
        vpt[4] -= e.deltaX;
        vpt[5] -= e.deltaY;
        this.canvas.setViewportTransform(vpt);
      }
      this.onChange({ reason: 'viewport', canUndo: this.canUndo, canRedo: this.canRedo });
    });
    // Finishes a paint gesture when a pointer is released outside the artboard.
    this._windowUp = () => this._pointerUp();
    window.addEventListener('pointerup', this._windowUp);
    window.addEventListener('blur', this._windowUp);
    this.newDocument({ width: 1200, height: 800, background: '#ffffff', name: '' });
  }

  get busy() { return this._busy; }
  get canUndo() { return this._historyIndex > 0; }
  get canRedo() { return this._historyIndex < this._history.length - 1; }
  get selected() { return this.canvas.getActiveObject() || null; }
  get layers() { return this.canvas.getObjects().filter((object) => !object.excludeFromExport); }
  get zoom() { return this.canvas.getZoom(); }
  _guard() { if (this._busy) throw new Error('BUSY'); }
  _emit(reason) {
    this.onChange({ reason, canUndo: this.canUndo, canRedo: this.canRedo });
  }
  _selectionChanged() {
    if(this._nodeSelectionId!==this.selected?.id)this.activeNodeIndex=null;
    this._nodeSelectionId=this.selected?.id;
    if (this.activeTool === 'node') this._enableNodeControls(this.selected);
    this.onSelection(this.selected);
  }
  _decorate(object, kind, name) {
    object.set({
      id: object.id || uid(), name: name || object.name || layerName(kind),
      cornerColor: '#e9fff8', cornerStrokeColor: '#187d73', borderColor: '#24b6a3',
      cornerSize: 9, transparentCorners: false, padding: 3,
    });
    this._applyLock(object);
    if (object instanceof Group) object.getObjects().forEach((child) => this._decorate(child, child.type));
    return object;
  }
  _applyLock(object) {
    object.set({
      selectable: !object.locked, evented: !object.locked,
      lockMovementX: !!object.locked, lockMovementY: !!object.locked,
      lockScalingX: !!object.locked, lockScalingY: !!object.locked,
      lockRotation: !!object.locked,
    });
  }
  _setArtboard() {
    this.canvas.backgroundColor = this.background === 'transparent' ? '' : this.background;
    this.canvas.clipPath = new Rect({
      left: 0, top: 0, width: this.width, height: this.height,
      originX: 'left', originY: 'top', absolutePositioned: true,
      fill: '#000', strokeWidth: 0, selectable: false, evented: false,
    });
    this.canvas.requestRenderAll();
  }

  newDocument({ width, height, background = '#ffffff', name = '', meta = {}, fonts = [] }) {
    this._guard();
    this._pointerUp();
    this._guard();
    width = Number(width); height = Number(height);
    validateDimensions(width, height);
    if (!CSS.supports('color', background)) throw new Error('INVALID_DOCUMENT');
    const normalizedMeta=normalizeMeta(meta,width,height);
    validateV2Extras({width,height,meta:normalizedMeta,fonts});
    this.cancelPath();
    this.canvas.discardActiveObject();
    this.canvas.clear();
    this.meta=normalizedMeta;this.fonts=structuredClone(fonts);this.activeNodeIndex=null;this.view.snapLines=[];
    this.width = width; this.height = height; this.background = background; this.name = String(name);
    this._setArtboard();
    this._history = []; this._historyIndex = -1; this._images.clear(); this._imageTokens.clear();
    this.activeTool = 'select'; this.setTool('select');
    this.commit('new');
  }

  serialize() {
    const canvas = this.canvas.toObject(CUSTOM);
    return {
      format: 'dake-stadio', version: 2,
      width: this.width, height: this.height, name: this.name, background: this.background,
      meta: structuredClone(this.meta), fonts: structuredClone(this.fonts),
      canvas: { version: canvas.version, objects: canvas.objects },
    };
  }

  async _withBusy(action) {
    this._guard();
    this._pointerUp();
    this._guard();
    this._busy = true;
    this.canvas.selection = false;
    this.canvas.skipTargetFind = true;
    this._emit('busy');
    try { return await action(); }
    finally {
      this._busy = false;
      this._applyTool();
      this._emit('ready');
    }
  }

  async _restore(data) {
    validateDocument(data);
    // Detached preflight owns every decoded object until all resources resolve.
    // A bad image can never clear the user's current canvas.
    const staging = new StaticCanvas(rasterCanvas(1, 1), { renderOnAddRemove: false });
    try {
      // Decode images without CPU filtering on the UI thread. Apply the same
      // serialized filter chain to detached objects in the worker before swap.
      const withoutFilters = (object) => ({
        ...object,
        ...(object.filters ? { filters: [] } : {}),
        ...(object.objects ? { objects: object.objects.map(withoutFilters) } : {}),
        ...(object.clipPath ? { clipPath: withoutFilters(object.clipPath) } : {}),
      });
      await staging.loadFromJSON({ objects: data.canvas.objects.map(withoutFilters) }, (_raw, object, error) => {
        if (error || !object) throw new Error('INVALID_DOCUMENT');
      });
      const all = staging.getObjects();
      if (all.length !== data.canvas.objects.length) throw new Error('INVALID_DOCUMENT');
      const check = async (object, raw) => {
        if (!object) throw new Error('INVALID_DOCUMENT');
        if (object instanceof FabricImage) {
          const image = object.getElement();
          validateDimensions(image.naturalWidth || image.width, image.naturalHeight || image.height);
          if (raw.filters?.length) await this._applyWorkerFilters(object, raw.filters);
          await this.restoreImageEditing?.(object,raw);
        }
        if (object instanceof Group) {
          if (object.getObjects().length !== (raw.objects?.length || 0)) throw new Error('INVALID_DOCUMENT');
          const children = object.getObjects();
          for (let index = 0; index < children.length; index++) await check(children[index], raw.objects[index]);
        }
        if (raw.clipPath) await check(object.clipPath, raw.clipPath);
        for (const key of ['fill', 'stroke']) if (raw[key] && typeof raw[key] === 'object' && !object[key]) throw new Error('INVALID_DOCUMENT');
      };
      for (let index = 0; index < all.length; index++) await check(all[index], data.canvas.objects[index]);
      staging.remove(...all);
      this.cancelPath();
      this.canvas.discardActiveObject();
      this.canvas.clear();
      this.width = data.width; this.height = data.height;
      this.background = data.background; this.name = data.name;
      this.meta=normalizeMeta(data.meta||{},data.width,data.height);this.fonts=structuredClone(data.fonts||[]);this.activeNodeIndex=null;this.view.snapLines=[];
      this.canvas.add(...all.map((object) => this._decorate(object, object.type)));
      this._setArtboard();
      this._revision++;
      this.onSelection(null);
    } finally { await staging.dispose(); }
  }

  async loadDocument(data, {prepare} = {}) {
    return this._withBusy(async () => {
      await prepare?.();
      await this._restore(data);
      this._history = []; this._historyIndex = -1; this._images.clear(); this._imageTokens.clear();
      this.commit('load');
    });
  }

  _snapshot() {
    return JSON.stringify(this.serialize(), (key, value) => {
      if ((['src','source'].includes(key) && typeof value === 'string' && IMAGE_DATA.test(value)) || (key === 'data' && typeof value === 'string' && /^data:(?:font\/|application\/)/i.test(value))) {
        let token = this._imageTokens.get(value);
        if (!token) {
          token = `@raster:${this._nextImageToken++}`;
          this._imageTokens.set(value, token); this._images.set(token, value);
        }
        return token;
      }
      return value;
    });
  }
  _expand(snapshot) {
    return JSON.parse(snapshot, (key, value) =>
      ['src','source','data'].includes(key) && typeof value === 'string' && value.startsWith('@raster:') ? this._images.get(value) : value);
  }
  _trimHistory() {
    const collect = () => {
      const used = new Set();
      for (const snapshot of this._history) for (const match of snapshot.matchAll(/@raster:\d+/g)) used.add(match[0]);
      for (const [token, source] of this._images) if (!used.has(token)) { this._images.delete(token); this._imageTokens.delete(source); }
      return this._history.reduce((total, item) => total + item.length * 2, 0) +
        Array.from(this._images.values()).reduce((total, item) => total + item.length * 2, 0);
    };
    let bytes = collect();
    while (this._history.length > 2 && (this._history.length > MAX_HISTORY || bytes > MAX_HISTORY_BYTES)) {
      this._history.shift(); this._historyIndex--; bytes = collect();
    }
  }
  commit(reason = 'edit') {
    const snapshot = this._snapshot();
    if (snapshot !== this._history[this._historyIndex]) {
      this._history.splice(this._historyIndex + 1);
      this._history.push(snapshot); this._historyIndex = this._history.length - 1;
      this._trimHistory(); this._revision++;
      this._emit(reason);
    }
    this.canvas.requestRenderAll();
    this.onSelection(this.selected);
  }
  async undo() {
    if (!this.canUndo) return false;
    return this._withBusy(async () => {
      const index = this._historyIndex - 1;
      await this._restore(this._expand(this._history[index]));
      this._historyIndex = index; this._emit('undo'); return true;
    });
  }
  async redo() {
    if (!this.canRedo) return false;
    return this._withBusy(async () => {
      const index = this._historyIndex + 1;
      await this._restore(this._expand(this._history[index]));
      this._historyIndex = index; this._emit('redo'); return true;
    });
  }

  _add(object, kind, name) {
    if (this.layers.length >= MAX_OBJECTS) throw new Error('INVALID_DOCUMENT');
    this._decorate(object, kind, name);
    this.canvas.add(object); this.canvas.setActiveObject(object);
    object.setCoords(); this.commit('add'); return object;
  }
  addShape(kind, options = {}) {
    this._guard();
    const size = Math.min(this.width, this.height) * 0.24;
    const props = {
      originX: 'left', originY: 'top', left: (this.width - size) / 2, top: (this.height - size) / 2,
      fill: this.style.fill, stroke: this.style.stroke, strokeWidth: this.style.strokeWidth,
      strokeUniform: true, ...options,
    };
    let object;
    if (kind === 'rect') object = new Rect({ width: size * 1.3, height: size, rx: 0, ry: 0, ...props });
    else if (kind === 'ellipse') object = new Ellipse({ rx: size * 0.65, ry: size * 0.5, ...props });
    else if (kind === 'triangle') object = new Triangle({ width: size * 1.2, height: size, ...props });
    else if (kind === 'line') object = new Line([0, 0, size * 1.5, 0], { ...props, stroke: this.style.stroke || this.style.fill, strokeWidth: this.style.strokeWidth || 3 });
    else throw new Error('INVALID_SHAPE');
    this.setTool('select');
    return this._add(object, kind);
  }
  addText(text = '') {
    this._guard(); this.setTool('select');
    const choice=this.fontSelection;
    if(choice){const descriptor=structuredClone(choice);if(descriptor.source==='local')delete descriptor.data;this.fonts=this.fonts.filter(f=>f.id!==descriptor.id);this.fonts.push(descriptor);}
    return this._add(new Textbox(String(text), {
      left: this.width * 0.2, top: this.height * 0.35, width: this.width * 0.6,
      originX: 'left', originY: 'top', fontFamily: choice?.family || 'Yu Gothic',
      fontWeight: choice?.weight || 400, fontStyle: choice?.style || 'normal',
      fontSize: Math.max(20, Math.round(this.width / 20)), fill: this.style.fill,
      stroke: this.style.stroke, strokeWidth: this.style.strokeWidth, lineHeight: 1.25,
    }), 'text');
  }

  async importImage(dataURL, name) {
    if (!IMAGE_DATA.test(dataURL) || dataURL.length > 128 * 1024 * 1024) throw new Error('INVALID_IMAGE');
    return this._withBusy(async () => {
      const image = await FabricImage.fromURL(dataURL);
      validateDimensions(image.width, image.height);
      const scale = Math.min(1, this.width * 0.88 / image.width, this.height * 0.88 / image.height);
      image.set({ originX: 'left', originY: 'top', scaleX: scale, scaleY: scale,
        left: (this.width - image.width * scale) / 2, top: (this.height - image.height * scale) / 2, objectCaching: false });
      this.activeTool = 'select';
      return this._add(image, 'image', name);
    });
  }
  async importSVG(svgString, name) {
    const safeSource = validateSVG(svgString);
    return this._withBusy(async () => {
      const result = await loadSVGFromString(safeSource);
      const objects = result.objects.filter(Boolean);
      if (!objects.length || objects.length > MAX_OBJECTS) throw new Error('UNSAFE_SVG');
      const object = objects.length === 1 ? objects[0] : new Group(objects, { originX: 'left', originY: 'top' });
      if (!Number.isFinite(object.width) || !Number.isFinite(object.height) || object.width < 0 || object.height < 0) throw new Error('UNSAFE_SVG');
      const scale = Math.min(1, this.width * 0.85 / Math.max(1, object.width), this.height * 0.85 / Math.max(1, object.height));
      object.set({ originX: 'left', originY: 'top', left: (this.width - object.width * scale) / 2,
        top: (this.height - object.height * scale) / 2, scaleX: scale, scaleY: scale });
      this.activeTool = 'select';
      return this._add(object, 'path', name);
    });
  }

  renameDocument(name) { this._guard(); this.name = String(name).slice(0, 1000); this.commit('rename'); }
  setBackground(color) {
    this._guard(); if (!CSS.supports('color', color)) return;
    this.background = color; this._setArtboard(); this.commit('background');
  }
  setStyle(style) {
    if (style.fill !== undefined) this.style.fill = style.fill;
    if (style.stroke !== undefined) this.style.stroke = style.stroke;
    if (style.strokeWidth !== undefined) this.style.strokeWidth = clamp(style.strokeWidth, 0, 200);
    if (style.brushSize !== undefined) this.style.brushSize = clamp(style.brushSize, 1, 1000);
    if (this._nodes.length) this._refreshPath();
  }
  updateSelected(props) {
    this._guard();
    const selected = this.canvas.getActiveObjects().filter((object) => !object.locked);
    if (!selected.length) return;
    const allowed = new Set(['left', 'top', 'angle', 'scaleX', 'scaleY', 'fill', 'stroke', 'strokeWidth', 'opacity', 'fontSize', 'fontFamily', 'fontWeight', 'fontStyle', 'underline', 'overline', 'linethrough', 'textAlign', 'charSpacing', 'lineHeight', 'text', 'name', 'flipX', 'flipY', 'rx', 'ry']);
    const sanitized = Object.fromEntries(Object.entries(props).filter(([key, value]) => allowed.has(key) && (typeof value !== 'number' || Number.isFinite(value))));
    if (sanitized.opacity !== undefined) sanitized.opacity = clamp(sanitized.opacity, 0, 1);
    if (sanitized.strokeWidth !== undefined) sanitized.strokeWidth = clamp(sanitized.strokeWidth, 0, 200);
    if (sanitized.charSpacing !== undefined) sanitized.charSpacing = clamp(sanitized.charSpacing, -500, 5000);
    if (sanitized.lineHeight !== undefined) sanitized.lineHeight = clamp(sanitized.lineHeight, 0.2, 10);
    if (sanitized.fontSize !== undefined) sanitized.fontSize = clamp(sanitized.fontSize, 1, 2000);
    if (sanitized.scaleX !== undefined) sanitized.scaleX = clamp(sanitized.scaleX, 0.001, 1000);
    if (sanitized.scaleY !== undefined) sanitized.scaleY = clamp(sanitized.scaleY, 0.001, 1000);
    for (const object of selected) { object.set(sanitized); object.setCoords(); }
    this.commit('properties');
  }
  nudge(dx, dy) {
    this._guard();
    const target = this.selected;
    if (!target || target.locked || target.isEditing) return;
    target.set({ left: target.left + dx, top: target.top + dy }); target.setCoords(); this.commit('nudge');
  }
  deleteSelected() {
    this._guard();
    const objects = this.canvas.getActiveObjects().filter((object) => !object.locked);
    if (!objects.length) return;
    this.canvas.discardActiveObject(); this.canvas.remove(...objects); this.commit('delete');
  }
  async duplicate() {
    if (!this.selected) return;
    return this._withBusy(async () => {
      const objects = this.canvas.getActiveObjects();
      this.canvas.discardActiveObject();
      const copies = await Promise.all(objects.map((object) => object.clone(CUSTOM)));
      const restoreCopy=async object=>{if(object instanceof FabricImage)await this.restoreImageEditing?.(object,object);if(object instanceof Group)for(const child of object.getObjects())await restoreCopy(child);};
      for(const object of copies)await restoreCopy(object);
      const reidentify = (object) => {
        object.id = uid(); object.locked = false;
        if(object.dakeMask?.source){object.dakeMask=structuredClone(object.dakeMask);object.dakeMask.source.id=uid();}
        if (object instanceof Group) object.getObjects().forEach(reidentify);
      };
      for (const object of copies) {
        reidentify(object); object.set({ left: object.left + 24, top: object.top + 24 });
        this.canvas.add(this._decorate(object, object.type));
      }
      this.canvas.setActiveObject(copies.length === 1 ? copies[0] : new ActiveSelection(copies, { canvas: this.canvas }));
      this.commit('duplicate'); return copies;
    });
  }
  group() {
    this._guard();
    const objects = this.canvas.getActiveObjects().filter(object=>!object.locked);
    if (objects.length < 2) { this.onStatus('groupNeedTwo'); return; }
    const indices = objects.map((object) => this.canvas.getObjects().indexOf(object));
    this.canvas.discardActiveObject(); this.canvas.remove(...objects);
    const group = this._decorate(new Group(objects, { originX: 'left', originY: 'top' }), 'group');
    this.canvas.insertAt(Math.min(...indices), group); this.canvas.setActiveObject(group); this.commit('group');
    return group;
  }
  ungroup() {
    this._guard();
    const group = this.selected;
    if (!(group instanceof Group) || group instanceof ActiveSelection || group.locked) return;
    if(group.dakeMask)return this.releaseMask();
    if(group.clipPath)throw new Error('INVALID_MASK');
    const index = this.canvas.getObjects().indexOf(group);
    this.canvas.discardActiveObject();
    const objects = group.removeAll();
    this.canvas.remove(group); this.canvas.insertAt(index, ...objects);
    objects.forEach((object) => { this._applyLock(object); object.setCoords(); });
    this.canvas.setActiveObject(new ActiveSelection(objects, { canvas: this.canvas })); this.commit('ungroup');
  }
  reorder(direction) {
    this._guard();
    const target = this.selected;
    if (!target) return;
    const methods = { up: 'bringObjectForward', down: 'sendObjectBackwards', top: 'bringObjectToFront', bottom: 'sendObjectToBack' };
    if (!methods[direction]) return;
    for (const object of this.canvas.getActiveObjects().filter(object=>!object.locked)) this.canvas[methods[direction]](object);
    this.commit('reorder');
  }
  select(id, {additive=false}={}) {
    this._guard();
    if(additive){const ids=this.canvas.getActiveObjects().map(object=>object.id);this.selectMany(ids.includes(id)?ids.filter(value=>value!==id):[...ids,id]);return this.selected;}
    const object = this.layers.find((layer) => layer.id === id);
    if (object) { this.canvas.setActiveObject(object); this.canvas.requestRenderAll(); }
    else this.canvas.discardActiveObject();
    return object;
  }
  toggleVisible(id) {
    this._guard(); const object = this.layers.find((layer) => layer.id === id);
    if (!object) return;
    object.set('visible', !object.visible); this.commit('visibility');
  }
  toggleLock(id) {
    this._guard(); const object = this.layers.find((layer) => layer.id === id);
    if (!object) return;
    object.locked = !object.locked; this._applyLock(object);
    if (this.selected === object) this.canvas.discardActiveObject();
    this.commit('lock');
  }

  fit(width, height) {
    const w = Math.max(1, Math.floor(width)); const h = Math.max(1, Math.floor(height));
    this.canvas.setDimensions({ width: w, height: h });
    const zoom = clamp(Math.min((w - 72) / this.width, (h - 72) / this.height), 0.02, 4);
    this.canvas.setViewportTransform([zoom, 0, 0, zoom, (w - this.width * zoom) / 2, (h - this.height * zoom) / 2]);
    this.canvas.requestRenderAll(); this._emit('viewport');
  }
  setZoom(zoom) {
    const center = new Point(this.canvas.width / 2, this.canvas.height / 2);
    this.canvas.zoomToPoint(center, clamp(zoom, 0.02, 16)); this._emit('viewport');
  }
  setTool(tool) {
    this._guard();
    if (!['select', 'pen', 'node', 'brush', 'eraser', 'hand', 'crop'].includes(tool)) return;
    if (tool !== this.activeTool) {
      this._pointerUp();
      this._guard();
      if (this.activeTool === 'pen' && this._nodes.length > 1) this.finishPath(false);
      else this.cancelPath();
    }
    this.activeTool = tool; this._applyTool(); this._emit('tool');
    if (tool === 'crop') this.onStatus('cropHint');
  }
  _applyTool() {
    const editing = ['select', 'node'].includes(this.activeTool) && !this._busy;
    this.canvas.selection = editing; this.canvas.skipTargetFind = !editing;
    this.canvas.defaultCursor = this.activeTool === 'hand' ? 'grab' : editing ? 'default' : 'crosshair';
    this.canvas.hoverCursor = editing ? 'move' : this.canvas.defaultCursor;
    for (const object of this.layers) {
      if (this._controls.has(object)) { object.controls = this._controls.get(object); object.hasBorders = true; }
    }
    if (this.activeTool === 'node') this._enableNodeControls(this.selected);
    this.canvas.requestRenderAll();
  }
  _enableNodeControls(object) {
    if (!(object instanceof Path) || object.locked) return;
    if (!this._controls.has(object)) this._controls.set(object, object.controls);
    const controls = controlsUtils.createPathControls(object, {
      pointStyle: { controlFill: '#ffffff', controlStroke: '#0b8b79' },
      controlPointStyle: { controlFill: '#ffc873', controlStroke: '#9b650c', connectionDashArray: [3, 3] },
    });
    // Keep a valid unaffected endpoint fixed. Fabric's default first-point
    // anchor reads the trailing Z command on a closed path, which has no point.
    for (const control of Object.values(controls)) {
      control.mouseDownHandler=()=>{this.activeNodeIndex=control.commandIndex;this._emit('node-selection');return true;};
      control.actionHandler = (_event, transform, x, y) => {
        const target = transform.target;
        const { commandIndex, pointIndex } = control;
        this.activeNodeIndex=commandIndex;
        const anchorIndex = target.path.findIndex((command, index) => command[0] !== 'Z' && index !== commandIndex);
        if (anchorIndex < 0) return false;
        const anchor = target.path[anchorIndex];
        const anchorPoint = new Point(anchor[anchor.length - 2], anchor[anchor.length - 1]);
        const fixed = anchorPoint.subtract(target.pathOffset).transform(target.calcTransformMatrix());
        const local = new Point(x, y).transform(util.invertTransform(target.calcTransformMatrix())).add(target.pathOffset);
        const command=target.path[commandIndex],oldX=command[pointIndex],oldY=command[pointIndex+1];
        if(command[0]==='M'&&pointIndex===1){
          let end=commandIndex+1;while(end<target.path.length&&target.path[end][0]!=='M')end++;
          const closure=target.path[end-2];
          if(target.path[end-1]?.[0]==='Z'&&closure&&closure.length>2&&closure.at(-2)===oldX&&closure.at(-1)===oldY){
            closure[closure.length-2]=local.x;closure[closure.length-1]=local.y;
            if(closure[0]==='C'){closure[3]+=local.x-oldX;closure[4]+=local.y-oldY;}
          }
        }
        command[pointIndex] = local.x; command[pointIndex + 1] = local.y;
        target.setDimensions();
        const moved = anchorPoint.subtract(target.pathOffset).transform(target.calcTransformMatrix());
        target.left += fixed.x - moved.x; target.top += fixed.y - moved.y;
        target.set('dirty', true); target.setCoords(); return true;
      };
    }
    object.controls = controls; object.hasBorders = false; object.setCoords();
  }

  _point(event) { return event.scenePoint || this.canvas.getScenePoint(event.e); }
  _pointerDown(event) {
    if (this._busy) {
      if (this._gestureTarget) this.canvas.setActiveObject(this._gestureTarget);
      return;
    }
    if (event.e.button !== undefined && event.e.button !== 0) return;
    const point = this._point(event);
    if (this.activeTool === 'hand') {
      this._drag = { kind: 'hand', x: event.e.clientX, y: event.e.clientY }; return;
    }
    if (this.activeTool === 'pen') {
      const first = this._nodes[0];
      if (first && this._nodes.length > 2 && Math.hypot(point.x - first.x, point.y - first.y) < 10 / this.zoom) { this.finishPath(true); return; }
      const last = this._nodes.at(-1);
      if (last && Math.hypot(point.x - last.x, point.y - last.y) < 2 / this.zoom) return;
      const node = { x: point.x, y: point.y, dx: 0, dy: 0 };
      this._nodes.push(node); this._pendingNode = node; this._refreshPath();
      this.onStatus('pathStarted'); return;
    }
    if (this.activeTool === 'brush' || this.activeTool === 'eraser') { this._startStroke(point); return; }
    if (this.activeTool === 'crop') {
      this._drag = { kind: 'crop', x: clamp(point.x, 0, this.width), y: clamp(point.y, 0, this.height), end: point };
      this._preview = new Rect({ left: point.x, top: point.y, width: 1, height: 1, originX: 'left', originY: 'top',
        fill: '#28b6a320', stroke: '#28b6a3', strokeWidth: 1 / this.zoom, strokeDashArray: [7 / this.zoom, 5 / this.zoom],
        selectable: false, evented: false, excludeFromExport: true });
      this.canvas.add(this._preview);
    }
  }
  _pointerMove(event) {
    if (this._busy) return;
    const point = this._point(event);
    if (this._drag?.kind === 'hand') {
      const vpt = this.canvas.viewportTransform.slice();
      vpt[4] += event.e.clientX - this._drag.x; vpt[5] += event.e.clientY - this._drag.y;
      this._drag.x = event.e.clientX; this._drag.y = event.e.clientY;
      this.canvas.setViewportTransform(vpt); return;
    }
    if (this._pendingNode) {
      this._pendingNode.dx = point.x - this._pendingNode.x; this._pendingNode.dy = point.y - this._pendingNode.y;
      this._refreshPath(); return;
    }
    if (this._stroke) { this._paintTo(point); return; }
    if (this._drag?.kind === 'crop' && this._preview) {
      const x = clamp(point.x, 0, this.width); const y = clamp(point.y, 0, this.height);
      this._drag.end = { x, y };
      this._preview.set({ left: Math.min(x, this._drag.x), top: Math.min(y, this._drag.y), width: Math.abs(x - this._drag.x), height: Math.abs(y - this._drag.y) });
      this.canvas.requestRenderAll();
    }
  }
  _pointerUp() {
    this._pendingNode = null;
    if (this._stroke) {
      const stroke = this._stroke; this._stroke = null;
      stroke.image.filters = [];
      stroke.image.setElement(stroke.buffer, { width: stroke.width, height: stroke.height });
      stroke.image.set('dirty', true); stroke.image.setCoords();
      if (stroke.filters.length) {
        this._drag = null;
        const task = this._withBusy(async () => {
          try {
            await this._applyWorkerFilters(stroke.image, stroke.filters.map((filter) => filter.toObject()));
            this.commit('raster');
            return true;
          } catch {
            // A worker failure restores both pixels and the visible filter
            // result. No incomplete stroke enters history or autosave.
            Object.assign(stroke.image, stroke.previous);
            stroke.image.set('dirty', true); stroke.image.setCoords();
            this.canvas.requestRenderAll(); this.onSelection(this.selected);
            this.onStatus('rasterFailed');
            return false;
          }
        });
        this._strokeTask = task;
        task.finally(() => { if (this._strokeTask === task) this._strokeTask = null; });
        return task;
      } else this.commit('raster');
    }
    if (this._drag?.kind === 'crop') {
      const { x, y, end } = this._drag;
      this._drag = null;
      if (this._preview) { this.canvas.remove(this._preview); this._preview = null; }
      const left = Math.round(Math.min(x, end.x)); const top = Math.round(Math.min(y, end.y));
      const width = Math.round(Math.abs(end.x - x)); const height = Math.round(Math.abs(end.y - y));
      if (width >= 2 && height >= 2) this.cropDocument({ left, top, width, height });
    }
    this._drag = null;
    return this._strokeTask;
  }
  _pathCommands(close = false) {
    const nodes = this._nodes;
    if (!nodes.length) return [];
    const commands = [['M', nodes[0].x, nodes[0].y]];
    const lineTo = (a, b) => commands.push(['C', a.x + a.dx, a.y + a.dy, b.x - b.dx, b.y - b.dy, b.x, b.y]);
    for (let i = 1; i < nodes.length; i++) lineTo(nodes[i - 1], nodes[i]);
    if (close) { lineTo(nodes.at(-1), nodes[0]); commands.push(['Z']); }
    return commands;
  }
  _refreshPath() {
    if (this._preview) this.canvas.remove(this._preview);
    const commands = this._pathCommands();
    if (commands.length === 1) commands.push(['L', this._nodes[0].x + 0.01, this._nodes[0].y + 0.01]);
    this._preview = new Path(commands, { fill: '', stroke: this.style.stroke || this.style.fill, strokeWidth: Math.max(this.style.strokeWidth, 2 / this.zoom),
      selectable: false, evented: false, excludeFromExport: true, objectCaching: false });
    this.canvas.add(this._preview); this.canvas.requestRenderAll();
  }
  finishPath(close = false) {
    if (this._nodes.length < 2) { this.cancelPath(); return null; }
    const path = new Path(this._pathCommands(close), { fill: close ? this.style.fill : '', stroke: this.style.stroke || this.style.fill,
      strokeWidth: this.style.strokeWidth || (close ? 0 : 3), strokeUniform: true, objectCaching: false });
    this.cancelPath(); const result = this._add(path, 'path'); this.onStatus('pathFinished'); return result;
  }
  cancelPath() {
    this._nodes = []; this._pendingNode = null;
    if (this._preview) { this.canvas.remove(this._preview); this._preview = null; }
    if (this._drag?.kind === 'crop') this._drag = null;
    this.canvas.requestRenderAll();
  }

  _startStroke(point) {
    let image = this._gestureTarget || this.selected;
    this._gestureTarget = null;
    if (!(image instanceof FabricImage) || image.locked) {
      if (this.activeTool === 'eraser') { this.onStatus('chooseImage'); return; }
      const buffer = rasterCanvas(this.width, this.height);
      image = this._decorate(new FabricImage(buffer, { left: 0, top: 0, originX: 'left', originY: 'top', objectCaching: false, dakeRaster: true }), 'raster');
      this.canvas.add(image); this.canvas.setActiveObject(image);
    }
    this.canvas.setActiveObject(image);
    const original = image._originalElement || image.getElement();
    const w = original.naturalWidth || original.width; const h = original.naturalHeight || original.height;
    const buffer = rasterCanvas(w, h); const context = buffer.getContext('2d', { willReadFrequently: true });
    context.drawImage(original, 0, 0);
    const oldFilters = image.filters;
    const width = image.width; const height = image.height;
    const previous = {
      _originalElement: image._originalElement, _element: image._element, _filteredEl: image._filteredEl,
      _filterScalingX: image._filterScalingX, _filterScalingY: image._filterScalingY,
      _lastScaleX: image._lastScaleX, _lastScaleY: image._lastScaleY,
      filters: oldFilters, adjustments: image.adjustments, dakeRaster: image.dakeRaster,
      objectCaching: image.objectCaching, width, height,
    };
    image.filters = [];
    image.setElement(buffer, { width, height });
    image.set({ dakeRaster: true, objectCaching: false });
    this._stroke = { image, buffer, context, filters: oldFilters, width, height, last: null, previous };
    this._paintTo(point);
  }
  _paintTo(scenePoint) {
    const { image, context, buffer } = this._stroke;
    const matrix = image.calcTransformMatrix();
    const point = new Point(scenePoint.x, scenePoint.y).transform(util.invertTransform(matrix));
    point.x += image.width / 2 + (image.cropX || 0); point.y += image.height / 2 + (image.cropY || 0);
    const scale = Math.sqrt(Math.abs(matrix[0] * matrix[3] - matrix[1] * matrix[2])) || 1;
    const radius = this.style.brushSize / scale / 2;
    context.globalCompositeOperation = this.activeTool === 'eraser' ? 'destination-out' : 'source-over';
    context.fillStyle = this.style.fill || '#000000'; context.strokeStyle = this.style.fill || '#000000';
    context.lineCap = 'round'; context.lineJoin = 'round'; context.lineWidth = radius * 2;
    const last = this._stroke.last;
    if (last) { context.beginPath(); context.moveTo(last.x, last.y); context.lineTo(point.x, point.y); context.stroke(); }
    else { context.beginPath(); context.arc(point.x, point.y, radius, 0, Math.PI * 2); context.fill(); }
    this._stroke.last = point;
    image._element = buffer; image.set('dirty', true);
    this.canvas.requestRenderAll();
  }
  _filterInWorker(bitmap, specifications) {
    if (this._workerJob) { bitmap.close(); return Promise.reject(new Error('BUSY')); }
    if (!this._worker) {
      this._worker = new Worker(new URL('./raster-worker.js', document.baseURI));
      this._worker.onmessage = ({ data }) => {
        const job = this._workerJob;
        if (!job || job.id !== data.id) { data.bitmap?.close(); return; }
        clearTimeout(job.timer); this._workerJob = null;
        if (data.error) job.reject(new Error(data.error)); else job.resolve(data.bitmap);
      };
      this._worker.onerror = () => {
        const job = this._workerJob; this._workerJob = null;
        if (job) { clearTimeout(job.timer); job.reject(new Error('INVALID_IMAGE')); }
        this._worker?.terminate(); this._worker = null;
      };
    }
    return new Promise((resolve, reject) => {
      const id = ++this._workerSequence;
      const timer = setTimeout(() => {
        this._worker?.terminate(); this._worker = null; this._workerJob = null;
        reject(new Error('INVALID_IMAGE'));
      }, 45_000);
      this._workerJob = { id, resolve, reject, timer };
      try { this._worker.postMessage({ id, bitmap, filters: specifications }, [bitmap]); }
      catch (error) { bitmap.close(); clearTimeout(timer); this._workerJob = null; reject(error); }
    });
  }
  async _applyWorkerFilters(image, specifications) {
    if (!specifications.length) {
      image.filters = []; image._element = image._originalElement;
      image._filteredEl = undefined; image._filterScalingX = 1; image._filterScalingY = 1;
      image.set('dirty', true); return;
    }
    const source = image._originalElement;
    const width = source.naturalWidth || source.width; const height = source.naturalHeight || source.height;
    validateDimensions(width, height);
    const bitmap = await createImageBitmap(source);
    const filteredBitmap = await this._filterInWorker(bitmap, specifications);
    try {
      const filtered = rasterCanvas(width, height);
      filtered.getContext('2d').drawImage(filteredBitmap, 0, 0);
      // Keep _originalElement untouched so source + Fabric filter definitions
      // remain fully editable and portable through project serialization.
      image.filters = specifications.map((specification) => new filters[specification.type](specification));
      image._filteredEl = filtered; image._element = filtered;
      image._filterScalingX = 1; image._filterScalingY = 1; image._lastScaleX = 1; image._lastScaleY = 1;
      image.set('dirty', true);
    } finally { filteredBitmap.close(); }
  }
  async applyImageAdjustment({ brightness = 0, contrast = 0, saturation = 0, grayscale = false } = {}) {
    this._guard();
    const image = this.selected;
    if (!(image instanceof FabricImage) || image.locked) throw new Error('NO_IMAGE_SELECTED');
    const settings = { brightness: clamp(brightness, -1, 1), contrast: clamp(contrast, -1, 1), saturation: clamp(saturation, -1, 1), grayscale: !!grayscale };
    return this._withBusy(async () => {
      const specifications = [];
      if (settings.brightness) specifications.push({ type: 'Brightness', brightness: settings.brightness });
      if (settings.contrast) specifications.push({ type: 'Contrast', contrast: settings.contrast });
      if (settings.saturation) specifications.push({ type: 'Saturation', saturation: settings.saturation });
      if (settings.grayscale) specifications.push({ type: 'Grayscale', mode: 'average' });
      await this._applyWorkerFilters(image, specifications);
      image.adjustments = settings; this.commit('adjustment');
    });
  }
  async rasterizeSelected() {
    const selected = this.selected;
    if (!selected || selected.locked) return;
    return this._withBusy(async () => {
      const rect = selected.getBoundingRect();
      validateDimensions(Math.max(1, Math.ceil(rect.width)), Math.max(1, Math.ceil(rect.height)));
      const buffer = selected.toCanvasElement({ multiplier: 1, enableRetinaScaling: false });
      const image = this._decorate(new FabricImage(buffer, {
        left: rect.left - (buffer.width - rect.width) / 2,
        top: rect.top - (buffer.height - rect.height) / 2,
        originX: 'left', originY: 'top', objectCaching: false, dakeRaster: true,
      }), 'raster', selected.name);
      const objects = this.canvas.getActiveObjects();
      const index = Math.min(...objects.map((object) => this.canvas.getObjects().indexOf(object)));
      this.canvas.discardActiveObject(); this.canvas.remove(...objects); this.canvas.insertAt(index, image);
      this.canvas.setActiveObject(image); this.commit('rasterize'); return image;
    });
  }
  resizeDocument(width,height,{scaleArtwork=false,anchor='center'}={}) {
    this._guard();width=Number(width);height=Number(height);validateDimensions(width,height);
    const sx=width/this.width,sy=height/this.height;
    if(scaleArtwork&&Math.abs(sx-sy)>.002)throw new Error('RESIZE_ASPECT');
    const dx=anchor==='center'?(width-this.width)/2:0,dy=anchor==='center'?(height-this.height)/2:0;
    this.canvas.discardActiveObject();
    for(const object of this.layers){object.set(scaleArtwork?{left:object.left*sx,top:object.top*sy,scaleX:object.scaleX*sx,scaleY:object.scaleY*sy}:{left:object.left+dx,top:object.top+dy});object.setCoords();}
    const limit=Math.max(0,(Math.min(width,height)-1)/2),bleed=Math.min(this.meta.bleed*(scaleArtwork?sx:1),limit),safe=Math.min(this.meta.safe*(scaleArtwork?sx:1),Math.max(0,limit-bleed));
    this.meta=normalizeMeta({...this.meta,bleed,safe,guides:{vertical:this.meta.guides.vertical.map(x=>scaleArtwork?x*sx:x+dx).filter(x=>x>=0&&x<=width),horizontal:this.meta.guides.horizontal.map(y=>scaleArtwork?y*sy:y+dy).filter(y=>y>=0&&y<=height)}},width,height);
    this.width=width;this.height=height;this._setArtboard();this.commit('resize-document');
  }
  cropDocument({ left, top, width, height }) {
    this._guard();
    left = Math.round(clamp(left, 0, this.width - 1)); top = Math.round(clamp(top, 0, this.height - 1));
    width = Math.round(clamp(width, 1, this.width - left)); height = Math.round(clamp(height, 1, this.height - top));
    validateDimensions(width, height);
    this.canvas.discardActiveObject();
    for (const object of this.layers) { object.set({ left: object.left - left, top: object.top - top }); object.setCoords(); }
    const limit=Math.max(0,(Math.min(width,height)-1)/2);
    this.meta=normalizeMeta({...this.meta,bleed:Math.min(this.meta.bleed,limit),safe:Math.min(this.meta.safe,Math.max(0,limit-this.meta.bleed)),guides:{vertical:this.meta.guides.vertical.map(x=>x-left).filter(x=>x>=0&&x<=width),horizontal:this.meta.guides.horizontal.map(y=>y-top).filter(y=>y>=0&&y<=height)}},width,height);
    this.width = width; this.height = height; this._setArtboard(); this.commit('crop'); this.onStatus('cropDone');
  }

  async exportRaster(format = 'png', scale = 1, quality = 0.94, {transparent=null} = {}) {
    this._guard(); this._pointerUp(); this._guard();
    if(this._exportPending)throw new Error('BUSY');
    if (!['png', 'jpeg', 'jpg', 'webp'].includes(format)) throw new Error('INVALID_IMAGE');
    scale = Number(scale);
    if (!Number.isFinite(scale) || scale <= 0 || this.width * this.height * scale * scale > 64_000_000 || Math.max(this.width, this.height) * scale > 16384) throw new Error('EXPORT_TOO_LARGE');
    const viewport = this.canvas.viewportTransform.slice();
    const type = format === 'jpg' ? 'jpeg' : format;
    let buffer;
    const savedBackground=this.canvas.backgroundColor;
    try {
      if(transparent===true)this.canvas.backgroundColor='';
      this.canvas.viewportTransform = identity();
      buffer = this.canvas.toCanvasElement(scale, { width: this.width, height: this.height, left: 0, top: 0, filter: (object) => !object.excludeFromExport });
      if (type === 'jpeg' || transparent===false) {
        const context = buffer.getContext('2d'); context.globalCompositeOperation = 'destination-over'; context.fillStyle = '#ffffff'; context.fillRect(0, 0, buffer.width, buffer.height);
      }
    } finally {
      this.canvas.backgroundColor=savedBackground;
      this.canvas.setViewportTransform(viewport); this.canvas.requestRenderAll();
    }
    // Encoding uses a detached immutable canvas. Restore the editing viewport
    // before yielding, and bound concurrent encodes to one allocation.
    this._exportPending=true;
    try {
      const blob=await new Promise((resolve,reject)=>buffer.toBlob(value=>value?resolve(value):reject(new Error('EXPORT_FAILED')),'image/'+type,clamp(quality,0.5,1)));
      return await new Promise((resolve,reject)=>{
        const reader=new FileReader();
        reader.onload=()=>typeof reader.result==='string'?resolve(reader.result):reject(new Error('EXPORT_FAILED'));
        reader.onerror=()=>reject(new Error('EXPORT_FAILED'));reader.onabort=reader.onerror;
        reader.readAsDataURL(blob);
      });
    } finally {
      buffer.width=buffer.height=1;this._exportPending=false;
    }
  }
  exportSVG() {
    this._guard(); this._pointerUp(); this._guard();
    const viewport = this.canvas.viewportTransform.slice();
    const width = this.canvas.width; const height = this.canvas.height;
    try {
      this.canvas.viewportTransform = identity(); this.canvas.width = this.width; this.canvas.height = this.height;
      return this.canvas.toSVG({ suppressPreamble: true, width: String(this.width), height: String(this.height), viewBox: { x: 0, y: 0, width: this.width, height: this.height } });
    } finally {
      this.canvas.width = width; this.canvas.height = height; this.canvas.setViewportTransform(viewport);
    }
  }
  async dispose() {
    window.removeEventListener('pointerup', this._windowUp); window.removeEventListener('blur', this._windowUp);
    if (this._workerJob) { clearTimeout(this._workerJob.timer); this._workerJob.reject(new Error('BUSY')); this._workerJob = null; }
    this._worker?.terminate(); this._worker = null;
    await this.canvas.dispose();
  }
}

installFoundation(StudioEngine);
installImaging(StudioEngine);
installComponents(StudioEngine);
