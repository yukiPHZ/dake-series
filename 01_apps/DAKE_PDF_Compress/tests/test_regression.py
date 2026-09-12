"""Quality rejection, original protection, candidate selection, CLI contracts."""
from pathlib import Path
import ast
import hashlib
import json
import multiprocessing
import os
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

APP=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(APP))
import main
import adaptive
import pymupdf as pdf
from benchmark import add_document_page


class Regression(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.folder=Path(self.temp.name)
        self.source=self.folder/'テスト document.pdf'
        with pdf.open() as d:
            for n in range(3): add_document_page(d,n+1)
            d.save(self.source)
        self.profile=adaptive.analyze(self.source)

    def tearDown(self):
        self.temp.cleanup()

    def test_structural_damage_rejected(self):
        for damage in ('missing','order','geometry','text'):
            path=self.folder/f'{damage}.pdf'
            with pdf.open(self.source) as d:
                if damage=='missing': d.delete_page(1)
                if damage=='order': d.select([2,1,0])
                if damage=='geometry': d[1].set_rotation(90)
                if damage=='text':
                    d[1].add_redact_annot(pdf.Rect(0,0,595,700));d[1].apply_redactions()
                d.save(path)
            with self.assertRaises(ValueError,msg=damage):
                adaptive.verify_candidate(self.source,path,self.profile,'B_balanced')

    def test_raster_stamp_and_lines_rejected(self):
        # No native text: only the visual checker can catch this loss.
        raster=self.folder/'raster.pdf'
        with pdf.open(self.source) as d, pdf.open() as r:
            pix=d[0].get_pixmap(dpi=144)
            p=r.new_page(width=595,height=842);p.insert_image(p.rect,pixmap=pix);r.save(raster)
        profile=adaptive.analyze(raster)
        for name,rect in [('stamp',pdf.Rect(460,700,535,775)),('lines',pdf.Rect(35,705,440,750))]:
            out=self.folder/f'{name}.pdf'
            with pdf.open(raster) as d:
                d[0].draw_rect(rect,color=None,fill=(1,1,1),overlay=True);d.save(out)
            with self.assertRaises(ValueError,msg=name):
                adaptive.verify_candidate(raster,out,profile,'C_document')

    def test_original_and_numbering(self):
        digest=hashlib.sha256(self.source.read_bytes()).hexdigest()
        old=self.source.with_name(self.source.stem+'_compressed.pdf');old.write_bytes(b'do not overwrite')
        a=main.compress_pdf(self.source);b=main.compress_pdf(self.source)
        self.assertTrue(a.output_path.name.endswith('_compressed_2.pdf'))
        self.assertTrue(b.output_path.name.endswith('_compressed_3.pdf'))
        self.assertEqual(old.read_bytes(),b'do not overwrite')
        self.assertEqual(digest,hashlib.sha256(self.source.read_bytes()).hexdigest())
        self.assertFalse(list(self.folder.glob('.dake_pdf_compress_*')))

    def test_selection_is_not_minimum_size(self):
        good={'name':'A_lossless','size':500,'checks':'passed','seconds':1,'worst':{'mae':0,'ink_loss':0}}
        degraded={'name':'B_balanced','size':490,'checks':'passed','seconds':5,'worst':{'mae':0.02,'ink_loss':0.01}}
        self.assertEqual(adaptive.choose_candidate([good,degraded],1000)['name'],'A_lossless')
        self.assertIsNone(adaptive.choose_candidate([dict(good,size=1000)],1000))

    def test_missing_rewrite_is_explicit_error(self):
        with patch.object(pdf.Document,'rewrite_images',None):
            with self.assertRaises(main.CompressError) as error: main.compress_pdf(self.source)
        self.assertEqual(error.exception.message_key,'error_dependency_missing')

    def test_ghostscript_optional(self):
        self.assertNotIn('D_ghostscript',adaptive.candidate_plan(self.profile))
        self.assertIn('D_ghostscript',adaptive.candidate_plan(self.profile,'fake.exe'))
        self.assertNotIn('D_ghostscript',adaptive.candidate_plan(dict(self.profile,interactive=True),'fake.exe'))
        with patch.object(main,'find_ghostscript',return_value=self.folder/'missing-gs.exe'):
            result=main.compress_pdf(self.source)
        self.assertEqual(result.engine,'A_lossless')

    def test_timeout(self):
        with self.assertRaises(TimeoutError):
            adaptive.run_candidate(self.source,self.folder/'out.pdf',self.profile,'A_lossless',None,lambda k:None,0.001)

    def test_ui_text(self):
        tree=ast.parse((APP/'main.py').read_text(encoding='utf-8'))
        missing=set()
        for node in ast.walk(tree):
            if isinstance(node,ast.Subscript) and isinstance(node.value,ast.Name) and node.value.id=='UI_TEXT' and isinstance(node.slice,ast.Constant):
                if node.slice.value not in main.UI_TEXT: missing.add(node.slice.value)
            if isinstance(node,ast.keyword) and node.arg=='text' and isinstance(node.value,ast.Constant) and isinstance(node.value.value,str):
                self.assertFalse(any(ord(c)>127 for c in node.value.value))
        self.assertFalse(missing)
        self.assertEqual(main.WINDOW_SIZE,'860x740');self.assertEqual(main.WINDOW_MIN_SIZE,(760,720))
        self.assertIn('シンプルそれDAKEシリーズ',main.UI_TEXT.values())


def cli_suite(command, folder):
    folder.mkdir(parents=True,exist_ok=True)
    one=folder/'正常 one.pdf';two=folder/'正常 two.pdf'
    for target in (one,two):
        with pdf.open() as d:
            add_document_page(d);d.save(target)
    invalid=folder/'broken.pdf';invalid.write_bytes(b'%PDF-1.7 broken')
    other=folder/'file.txt';other.write_text('not a pdf')
    empty=folder/'compact.pdf'
    with pdf.open() as d:
        d.new_page();d.save(empty,garbage=4,deflate=True,use_objstms=1,clean=True)
    cases=[('help',['--help-cli'],0),('one',['--from-shimarisu','--inputs',str(one)],0),
           ('many',['--from-shimarisu','--inputs',str(one),str(two)],0),
           ('missing',['--from-shimarisu','--inputs',str(folder/'missing.pdf')],1),
           ('not_pdf',['--from-shimarisu','--inputs',str(other)],1),
           ('broken',['--from-shimarisu','--inputs',str(invalid)],1),
           ('no_inputs',['--from-shimarisu'],1),
           ('no_reduction',['--from-shimarisu','--inputs',str(empty)],1)]
    results=[]
    for name,args,code in cases:
        start=time.monotonic()
        env=dict(os.environ,DAKE_PDF_COMPRESS_DEBUG='1',PYTHONIOENCODING='utf-8')
        p=subprocess.run(command+args,capture_output=True,timeout=180,env=env)
        out=p.stdout.decode('utf-8');err=p.stderr.decode('utf-8')
        assert p.returncode==code,(name,p.returncode,out,err)
        assert 'Traceback' not in out+err,(name,out,err)
        if code==0:
            assert not err,(name,err)
            if name!='help':
                paths=out.strip().splitlines()
                assert len(paths)==(2 if name=='many' else 1)
                assert all(Path(x).is_file() for x in paths),(name,paths)
        else:
            assert not out and err.strip() and len(err.strip().splitlines())==1,(name,out,err)
        results.append({'case':name,'exit':p.returncode,'stdout':out,'stderr':err,'seconds':time.monotonic()-start})
    return results


if __name__=='__main__':
    multiprocessing.freeze_support()
    unittest.main()
