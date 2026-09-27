import paper from 'paper/dist/paper-core.js';
import { parse } from 'opentype.js';

const scope = new paper.PaperScope();
scope.setup(new scope.Size(1, 1));
const pathString = (commands) => commands.map((command) => command.join(' ')).join(' ');
function geometry(item) {
  let path;
  const w = item.width, h = item.height;
  if (item.children) {
    const children = item.children.map(geometry);
    path = children.shift();
    for (const child of children) { const next = path.unite(child, {insert:false}); path.remove(); child.remove(); path = next; }
    return path;
  }
  if (item.type === 'path') {
    path = scope.PathItem.create(pathString(item.path));
    path.translate(new scope.Point(-item.pathOffset.x, -item.pathOffset.y));
  } else if (item.type === 'rect') {
    path = new scope.Path.Rectangle({rectangle:new scope.Rectangle(-w/2,-h/2,w,h), radius:new scope.Size(Math.min(item.rx||0,w/2),Math.min(item.ry||0,h/2)),insert:false});
  } else if (item.type === 'ellipse' || item.type === 'circle') {
    path = new scope.Path.Ellipse({rectangle:new scope.Rectangle(-w/2,-h/2,w,h),insert:false});
  } else if (item.type === 'triangle') {
    path = new scope.Path({segments:[[-w/2,h/2],[0,-h/2],[w/2,h/2]],closed:true,insert:false});
  } else if (item.type === 'polygon' || item.type === 'polyline') {
    path = new scope.Path({segments:item.points.map(p=>[p.x-item.pathOffset.x,p.y-item.pathOffset.y]),closed:item.type==='polygon',insert:false});
  } else if (item.type === 'line') {
    path = new scope.Path({segments:[[item.x1,item.y1],[item.x2,item.y2]],insert:false});
  } else throw new Error('NO_VECTOR_SELECTED');
  path.fillRule = item.fillRule || 'nonzero';
  path.transform(new scope.Matrix(item.matrix));
  return path;
}
function contours(path) {
  return path.className === 'CompoundPath' ? path.children.slice() : [path];
}
function svgPaths(item, separate=false) {
  if (item.className === 'Group' || separate && item.className === 'CompoundPath') return item.children.flatMap(child=>svgPaths(child,false));
  return item.pathData && item.pathData.length ? [item.pathData] : [];
}
function booleanJob(payload) {
  const items = payload.items.map(geometry);
  const operation = payload.operation;
  if (operation === 'convert') return {paths:items.flatMap(item=>svgPaths(item))};
  if (operation === 'split') return {paths:items.flatMap(item=>svgPaths(item,true))};
  if (operation === 'compound') {
    const compound = new scope.CompoundPath({children:items.flatMap(contours).map(path=>path.clone({insert:false})),insert:false});
    compound.fillRule = 'evenodd';
    return {paths:[compound.pathData],fillRule:'evenodd'};
  }
  if (items.length < 2) throw new Error('NEED_TWO_SHAPES');
  for (const item of items) for (const contour of contours(item)) contour.closed = true;
  if (operation === 'divide') {
    let pieces = [items[0]];
    for (const cutter of items.slice(1)) {
      pieces = pieces.flatMap(piece=> {
        const outside = piece.subtract(cutter,{insert:false});
        const inside = piece.intersect(cutter,{insert:false});
        return [outside,inside].filter(part=>Math.abs(part.area)>0.000001);
      });
      if (pieces.length > 2000) throw new Error('VECTOR_TOO_COMPLEX');
    }
    return {paths:pieces.flatMap(path=>svgPaths(path)),fillRule:'nonzero'};
  }
  const methods = {union:'unite',subtract:'subtract',intersect:'intersect'};
  if (!methods[operation]) throw new Error('NO_VECTOR_SELECTED');
  let result = items[0];
  for (const other of items.slice(1)) result = result[methods[operation]](other,{insert:false});
  return {paths:svgPaths(result),fillRule:result.fillRule||'nonzero'};
}
function outlineJob(payload) {
  const font = parse(payload.font);
  if (font.tables.colr || font.tables.svg) throw new Error('FONT_UNSUPPORTED');
  const axes = font.tables.fvar?.axes || [];
  const variation = {};
  if(!axes.some(axis=>axis.tag==='wght')&&Math.abs((font.tables.os2?.usWeightClass||400)-payload.weight)>100)throw new Error('FONT_UNSUPPORTED');
  for (const axis of axes) {
    const wanted = axis.tag==='wght' ? payload.weight : axis.defaultValue;
    variation[axis.tag]=Math.max(axis.minValue,Math.min(axis.maxValue,wanted));
  }
  if (axes.length && font.variation) font.variation.set(variation);
  if (payload.style !== 'normal' && !/italic|oblique/i.test(font.names.fontSubfamily?.en||'') && !axes.some(a=>a.tag==='ital'||a.tag==='slnt')) throw new Error('FONT_UNSUPPORTED');
  if (axes.some(a=>a.tag==='ital') && payload.style==='italic') variation.ital=1;
  if (Object.keys(variation).length) font.variation.set(variation);
  const commands=[];
  const matrix=payload.matrix;
  const pt=(x,y)=>[matrix[0]*x+matrix[2]*y+matrix[4],matrix[1]*x+matrix[3]*y+matrix[5]];
  for (const run of payload.runs) {
    for (const char of Array.from(run.text)) if (!/\s/u.test(char) && !font.hasChar(char)) throw new Error('FONT_GLYPH_MISSING');
    const glyphPaths=[];
    if(run.positions){
      // The canvas owns text layout. Its measured pair positions also cover
      // GPOS lookups that OpenType.js does not implement (e.g. Arial kerning).
      // Preserve recognized ligatures; reject unrecognized shaping explicitly.
      const glyphs=font.stringToGlyphs(run.text);
      let index=0;
      for(const glyph of glyphs){
        let count=1;
        if(glyph.index!==font.charToGlyphIndex(run.chars[index]||'')){
          const expanded=glyph.unicode?String.fromCodePoint(glyph.unicode).normalize('NFKD'):'';
          if(!expanded||!run.chars.slice(index).join('').startsWith(expanded))throw new Error('FONT_UNSUPPORTED');
          count=Array.from(expanded).length;
        }
        glyphPaths.push(glyph.getPath(run.x+(run.positions[index]||0),run.y,payload.fontSize,{variation},font));index+=count;
      }
      if(index!==run.chars.length)throw new Error('FONT_UNSUPPORTED');
    }else glyphPaths.push(font.getPath(run.text,run.x,run.y,payload.fontSize,{kerning:true,features:{liga:false,rlig:true},variation}));
    for(const glyphPath of glyphPaths)for (const command of glyphPath.commands) {
      if(command.type==='Z')commands.push(['Z']);
      else if(command.type==='M'||command.type==='L')commands.push([command.type,...pt(command.x,command.y)]);
      else if(command.type==='Q')commands.push(['Q',...pt(command.x1,command.y1),...pt(command.x,command.y)]);
      else if(command.type==='C')commands.push(['C',...pt(command.x1,command.y1),...pt(command.x2,command.y2),...pt(command.x,command.y)]);
    }
    if (commands.length>100000) throw new Error('VECTOR_TOO_COMPLEX');
  }
  if(!commands.length)throw new Error('VECTOR_EMPTY');
  let data=pathString(commands);
  if(payload.decorations?.length){
    let geometry=scope.PathItem.create(data);
    for(const rect of payload.decorations){
      const box=new scope.Path.Rectangle({rectangle:new scope.Rectangle(rect.x,rect.y,rect.width,rect.height),insert:false});
      box.transform(new scope.Matrix(matrix));geometry=geometry.unite(box,{insert:false});
    }
    data=geometry.pathData;
  }
  return {paths:[data],fillRule:'nonzero',variation};
}
self.onmessage=({data:{id,method,payload}})=>{
  try {
    scope.activate(); scope.project.clear();
    const result=method==='outline'?outlineJob(payload):booleanJob(payload);
    if(result.paths.join('').length>8*1024*1024)throw new Error('VECTOR_TOO_COMPLEX');
    self.postMessage({id,result});
  } catch(error) { self.postMessage({id,error:error.message||'VECTOR_TOO_COMPLEX'}); }
  finally { scope.project.clear(); }
};

