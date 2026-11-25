from pathlib import Path
import importlib.util
import json
import multiprocessing
import os
import sys

ROOT=Path(__file__).resolve().parent
APP=ROOT/'01_apps/DAKE_PDF_Compress'
sys.path.insert(0,str(APP/'tests'))
sys.path.insert(0,str(APP))
from test_regression import cli_suite
from benchmark import run

if __name__=='__main__':
    multiprocessing.freeze_support()
    if sys.argv[1]=='cli':
        source=cli_suite([sys.executable,str(APP/'main.py')],ROOT/'evidence/cli-source')
        exe=cli_suite([str(APP/'dist/DakePDF_Compress.exe')],ROOT/'evidence/cli-exe')
        (ROOT/'evidence/cli.json').write_text(json.dumps({'source':source,'exe':exe},ensure_ascii=False,indent=2),encoding='utf-8')
        print('source/exe CLI:',len(source),len(exe),'PASS')
    elif sys.argv[1]=='visual':
        import pymupdf as p
        import adaptive
        rows=json.loads((ROOT/'evidence/benchmark.json').read_text(encoding='utf-8'))
        dest=ROOT/'evidence/visual';dest.mkdir(exist_ok=True)
        for r in rows:
            src=ROOT/'fixtures'/r['file']
            chosen=r['selected']
            if not chosen:continue
            for label,path in [('original',src),('selected',chosen['path']),('legacy',ROOT/'evidence'/(src.stem+'_old.pdf'))]:
                if not Path(path).exists():continue
                with p.open(path) as d:
                    d[0].get_pixmap(dpi=100).save(dest/f'{src.stem}_{label}.png')
                    if src.stem[0] in 'CDE':
                        d[0].get_pixmap(dpi=250,clip=p.Rect(25,680,545,780)).save(dest/f'{src.stem}_{label}_detail.png')
        # Explicit FAX result is shown even if A is selected for smaller size.
        src=ROOT/'fixtures/C_bitonal.pdf';out=ROOT/'evidence/C_fax_candidate.pdf'
        adaptive.create_candidate(src,out,'C_bitonal_fax')
        with p.open(src) as a,p.open(out) as b:
            # Pixel-exact at native scan DPI, all three fixture pages.
            assert all(a[i].get_pixmap(dpi=300).samples==b[i].get_pixmap(dpi=300).samples for i in range(len(a)))
            b[0].get_pixmap(dpi=250,clip=p.Rect(25,680,545,780)).save(dest/'C_bitonal_FAX_detail.png')
        print('visual evidence generated; FAX 300-dpi all-page equality PASS')
