import sys
sys.dont_write_bytecode = True
import json
from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageChops, ImageStat
root=Path(__file__).resolve().parent.parent
out=root/'test-output/native-pdf-v2'
report=[]
for name,dpi in [('business-card-bleed',300),('a5-landscape',212),('transparent-composite',144)]:
    pdf=pdfium.PdfDocument(str(out/(name+'.pdf')))
    assert len(pdf)==1
    page=pdf[0]
    rendered=page.render(scale=dpi/72*(1-1e-7)).to_pil().convert('RGB')
    source=Image.open(out/(name+'.png')).convert('RGBA')
    original=Image.alpha_composite(Image.new('RGBA',source.size,'white'),source).convert('RGB')
    rendered.save(out/(name+'-rendered.png'))
    assert rendered.size==original.size,(rendered.size,original.size)
    difference=ImageChops.difference(rendered,original)
    rms=ImageStat.Stat(difference).rms
    assert max(rms)<1.0,rms
    report.append({'name':name,'size':list(rendered.size),'rmsRgb':rms,'passed':True})
    page.close();pdf.close()
(root/'evidence/native-pdf-render-results.json').write_text(json.dumps({'passed':True,'renders':report},indent=2),encoding='utf8')
print(json.dumps(report,indent=2))
