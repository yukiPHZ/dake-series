"""Bounded, latest-request workers. All filesystem/decoder work stays off Tk."""
from collections import OrderedDict
from pathlib import Path
import os
import threading
import queue
import types
from PIL import Image, ImageOps
import pypdfium2 as pdfium
from rename_core import Snapshot, plain_path

class LatestWorker:
    def __init__(self, function):
        self.function = function
        self.condition = threading.Condition()
        self.pending = None
        self.closed = False
        self.generation = 0
        self.output = queue.Queue(maxsize=1)
        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()
    def submit(self, value):
        with self.condition:
            self.generation += 1
            self.pending = (self.generation, value)
            self.condition.notify()
            return self.generation
    def run(self):
        while True:
            with self.condition:
                self.condition.wait_for(lambda: self.closed or self.pending is not None)
                if self.closed:
                    return
                generation, value = self.pending
                self.pending = None
            try:
                result = self.function(value, lambda: self.closed or generation != self.generation)
                if isinstance(result, types.GeneratorType):
                    for partial in result:
                        if generation != self.generation or self.closed: break
                        self.publish((generation,partial,None))
                    continue
                error = None
            except Exception as exc:
                result, error = None, str(exc)
            if generation == self.generation and not self.closed:
                self.publish((generation,result,error))
    def publish(self, value):
        try: self.output.get_nowait()
        except queue.Empty: pass
        self.output.put_nowait(value)
    def close(self):
        with self.condition:
            self.closed = True
            self.pending = None
            self.condition.notify_all()

def scan(request, cancelled):
    paths, parent = request
    candidates, folders, failures = [], [], []
    for raw in paths:
        if cancelled(): return [], []
        try:
            path = plain_path(raw)
            if parent:
                with os.scandir(path) as entries:
                    candidates.extend(Path(e.path) for e in entries if e.is_dir(follow_symlinks=False) or e.is_symlink())
            else:
                candidates.append(path)
        except Exception as exc:
            failures.append((str(raw), str(exc)))
    seen = set()
    for path in sorted(candidates, key=lambda p:p.name.casefold()):
        if cancelled(): return [], []
        try:
            key = str(path).casefold()
            if key in seen: continue
            seen.add(key)
            snapshot = Snapshot.capture(path)
            files, subdirs = [], 0
            with os.scandir(path) as entries:
                for entry in entries:
                    if cancelled(): return [], []
                    try:
                        s = entry.stat(follow_symlinks=False)
                        link = entry.is_symlink() or bool(getattr(s,'st_file_attributes',0)&0x400)
                        directory = entry.is_dir(follow_symlinks=False)
                        if directory: subdirs += 1
                        files.append({'path':Path(entry.path),'name':entry.name,'size':s.st_size,'mtime':s.st_mtime_ns,'directory':directory,'link':link,'error':''})
                    except OSError as exc:
                        files.append({'path':Path(entry.path),'name':entry.name,'size':None,'mtime':0,'directory':False,'link':True,'error':str(exc)})
            files.sort(key=lambda f:f['name'].casefold())
            folders.append({'snapshot':snapshot,'files':files,'subdirs':subdirs,'proposed':'','checked':True})
        except Exception as exc:
            failures.append((str(path),str(exc)))
    return folders, failures

class Renderer:
    def __init__(self):
        self.cache = OrderedDict()
    def __call__(self, files, cancelled):
        result=[]
        for index, item in files:
            if cancelled(): break
            key = (str(item['path']),item['size'],item['mtime'])
            if key not in self.cache:
                self.cache[key] = self.render(item)
                while len(self.cache)>64:
                    self.cache.popitem(last=False)
            self.cache.move_to_end(key)
            result.append((index,self.cache[key]))
            yield list(result)
    def render(self,item):
        suffix=item['path'].suffix.lower()
        if item['error']: return None,None,'unreadable'
        if item['link']: return None,None,'link'
        if item['directory']: return None,None,'folder'
        try:
            plain_path(item['path'])
            if suffix=='.pdf':
                with pdfium.PdfDocument(str(item['path'])) as doc:
                    pages=len(doc)
                    page=doc[0]
                    try:
                        width,height=page.get_size()
                        bitmap=page.render(scale=min(300/max(width,1),220/max(height,1)))
                        try: im=bitmap.to_pil().convert('RGB').copy()
                        finally: bitmap.close()
                    finally: page.close()
                return im,pages,'pdf'
            if suffix in ('.jpg','.jpeg','.png','.webp','.bmp','.tif','.tiff'):
                with Image.open(item['path']) as source:
                    source.thumbnail((300,220))
                    im=ImageOps.exif_transpose(source).convert('RGB').copy()
                return im,None,'image'
            return None,None,'other'
        except Exception:
            return None,None,'preview_failed'
