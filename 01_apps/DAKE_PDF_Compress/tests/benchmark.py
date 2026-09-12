"""Deterministic, fictional fixtures and candidate/legacy comparison. Dev only."""
from __future__ import annotations
import argparse
import hashlib
import json
import math
from pathlib import Path
import random
import sys
import time
import zlib

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
import adaptive
import pymupdf as pdf


def add_document_page(doc, number=1):
    p = doc.new_page(width=595, height=842)
    p.insert_text((40, 48), "DAKE TEST / FICTIONAL DOCUMENT", fontsize=18)
    p.insert_text((40, 70), "No personal or customer information. SAMPLE ONLY.", fontsize=10)
    for i in range(42):
        p.insert_text((42, 96+i*14), f"{number:03d}-{i:02d}  Contract sample 0123456789  Il1 O0  rn m  8B  [ ]", fontsize=7 if i%3 else 5)
    for i, width in enumerate((0.15, 0.25, 0.5, 0.75, 1)):
        p.draw_line((40, 710+i*8), (430, 710+i*8), width=width)
    p.insert_text((40, 680), "検証専用・架空文書　契約書／印影／図面　小さい文字 123456", fontname="japan", fontsize=8)
    p.draw_circle((495, 735), 28, color=(0.75,0.08,0.12), width=0.6)
    p.insert_text((478,738), "TEST", fontsize=11, color=(0.75,0.08,0.12))
    p.insert_text((40, 805), f"PAGE {number:03d}", fontsize=10)
    return p


def raster_document(bitonal=False, gray=False):
    with pdf.open() as doc:
        p = add_document_page(doc)
        pix = p.get_pixmap(dpi=300, colorspace=pdf.csGRAY if bitonal or gray else pdf.csRGB)
        if not bitonal:
            return pix, None
        # Pack true 1-bit pixels without an imaging dependency.
        data = pix.samples
        stride = (pix.width + 7)//8
        packed = bytearray([255]) * (stride * pix.height)
        for y in range(pix.height):
            for x in range(pix.width):
                if data[y*pix.width+x] < 180:
                    packed[y*stride+x//8] &= ~(128 >> (x%8))
        return pix, bytes(packed)


def put_bitonal(doc, page, pix, packed):
    xref = doc.get_new_xref()
    doc.update_object(xref, f"<< /Type /XObject /Subtype /Image /Width {pix.width} /Height {pix.height} /BitsPerComponent 1 /ColorSpace /DeviceGray >>")
    doc.update_stream(xref, packed, compress=False)
    resources = doc.get_new_xref()
    doc.update_object(resources, f"<< /XObject << /Scan {xref} 0 R >> >>")
    doc.xref_set_key(page.xref, "Resources", f"{resources} 0 R")
    stream = doc.get_new_xref()
    doc.update_object(stream, "<<>>")
    doc.update_stream(stream, b"q 595 0 0 842 0 0 cm /Scan Do Q", compress=False)
    page.set_contents(stream)


def photo_pixmap():
    # Procedural landscape with continuous tones, texture, and hard detail.
    # Synthetic photographic-like stress fixture, explicitly not a real photo.
    width, height = 1800, 1200
    rng = random.Random(812)
    values = bytearray(width*height*3)
    for y in range(height):
        for x in range(width):
            horizon = 580 + 90*math.sin(x/250) + 40*math.sin(x/70)
            noise = rng.randrange(-10,11)
            if y < horizon:
                rgb = (85+70*y/height, 135+75*y/height, 210+30*y/height)
            else:
                rgb = (40+50*y/height, 75+65*y/height, 35+35*y/height)
            if (x-1300)**2+(y-240)**2 < 130**2:
                rgb = (242,223,154)
            i=(y*width+x)*3
            values[i:i+3]=bytes(max(0,min(255,int(c)+noise)) for c in rgb)
    return pdf.Pixmap(pdf.csRGB,width,height,bytes(values),False)


def fixtures(folder):
    folder.mkdir(parents=True,exist_ok=True)
    with pdf.open() as d:
        for n in range(1,4): add_document_page(d,n)
        d.save(folder/'A_text.pdf')
    with pdf.open() as d:
        pix=photo_pixmap()
        for n in range(3):
            p=d.new_page(width=595,height=420)
            p.insert_image(p.rect,pixmap=pix)
            p.insert_text((30,400),f"SYNTHETIC LANDSCAPE {n+1}",fontsize=12)
        d.save(folder/'B_photo_like.pdf')
    bitpix,bits=raster_document(bitonal=True)
    scanpix,_=raster_document()
    graypix,_=raster_document(gray=True)
    for code,count,kind in [('C_bitonal',3,'bit'),('D_color_scan',3,'color'),('E_fine_drawing',3,'gray'),('G_60_pages',60,'bit'),('H_100_pages',100,'color')]:
        with pdf.open() as d:
            for n in range(count):
                p=d.new_page(width=595,height=842)
                if kind=='bit': put_bitonal(d,p,bitpix,bits)
                else: p.insert_image(p.rect,pixmap=scanpix if kind=='color' else graypix)
                # Native unique folio detects missing or reordered pages.
                p.insert_text((450,822),f"TEST PAGE {n+1:03d}",fontsize=6)
            d.save(folder/f'{code}.pdf')
    adaptive.create_candidate(folder/'B_photo_like.pdf',folder/'F_precompressed.pdf','B_balanced')


def run(folder, output, baseline=None):
    rows=[]
    for source in sorted(folder.glob('*.pdf')):
        if '_compressed' in source.stem: continue
        digest=hashlib.sha256(source.read_bytes()).hexdigest()
        start=time.monotonic()
        selected,profile,candidates=adaptive.compress(source)
        row={'file':source.name,'original':source.stat().st_size,'seconds':time.monotonic()-start,
             'selected':selected,'profile':{k:v for k,v in profile.items() if k!='page_info'},'candidates':candidates}
        if baseline:
            import importlib.util
            spec=importlib.util.spec_from_file_location('legacy',baseline)
            old=importlib.util.module_from_spec(spec);sys.modules['legacy']=old;spec.loader.exec_module(old)
            oldfile=output.parent/(source.stem+'_old.pdf')
            start=time.monotonic()
            old.save_pymupdf_compressed(source,oldfile)
            row['legacy']={'size':oldfile.stat().st_size,'seconds':time.monotonic()-start}
            try: row['legacy']['validation']=adaptive.verify_candidate(source,oldfile,profile,'B_balanced')
            except Exception as e: row['legacy']['rejected']=str(e)
        assert digest==hashlib.sha256(source.read_bytes()).hexdigest()
        rows.append(row)
        output.write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
        print(source.name,profile['kind'],source.stat().st_size,selected and selected['size'],round(row['seconds'],2),flush=True)
    return rows


if __name__=='__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    ap=argparse.ArgumentParser();ap.add_argument('--fixtures',type=Path,required=True);ap.add_argument('--results',type=Path);ap.add_argument('--baseline',type=Path);ap.add_argument('--generate',action='store_true');args=ap.parse_args()
    if args.generate: fixtures(args.fixtures)
    if args.results: run(args.fixtures,args.results,args.baseline)
