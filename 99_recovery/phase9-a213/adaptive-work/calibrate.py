from pathlib import Path
import json
import sys
ROOT=Path(__file__).parent
sys.path.insert(0,str(ROOT/'01_apps/DAKE_PDF_Compress'))
import adaptive
import pymupdf as pdf

rows=[]
source=ROOT/'fixtures/D_color_scan.pdf'
profile=adaptive.analyze(source)
for name,rect in [('stamp',pdf.Rect(460,700,535,775)),('fine_lines',pdf.Rect(35,705,440,750))]:
    path=ROOT/'evidence'/f'bad_{name}.pdf'
    with pdf.open(source) as d:
        d[0].draw_rect(rect,color=None,fill=(1,1,1),overlay=True);d.save(path)
    with pdf.open(source) as a,pdf.open(path) as b:
        metrics=adaptive.visual_difference(adaptive.render(a[0]),adaptive.render(b[0]))
    try:
        adaptive.verify_candidate(source,path,profile,'C_document')
        accepted=True
    except ValueError:accepted=False
    assert not accepted,name
    rows.append({'control':name,'metrics':metrics,'accepted':accepted})
source=ROOT/'fixtures/B_photo_like.pdf';path=ROOT/'evidence/bad_photo.pdf'
with pdf.open(source) as d:
    d.rewrite_images(dpi_threshold=100,dpi_target=36,quality=1,bitonal=False,gray=False)
    d.save(path,garbage=4,deflate=True)
with pdf.open(source) as a,pdf.open(path) as b:
    metrics=adaptive.visual_difference(adaptive.render(a[0]),adaptive.render(b[0]))
try:
    adaptive.verify_candidate(source,path,adaptive.analyze(source),'B_balanced')
    accepted=True
except ValueError:accepted=False
assert not accepted,metrics
rows.append({'control':'36dpi JPEG1 extreme degradation','metrics':metrics,'accepted':accepted})
(ROOT/'evidence/calibration.json').write_text(json.dumps({'limits':adaptive.QUALITY_LIMITS,'controls':rows},indent=2),encoding='utf-8')
print(rows)
