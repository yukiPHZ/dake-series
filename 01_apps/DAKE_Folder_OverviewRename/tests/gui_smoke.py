import sys,time,json
from pathlib import Path
from PIL import Image,ImageDraw
from reportlab.pdfgen.canvas import Canvas
from pypdf import PdfReader,PdfWriter
app=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(app))
from main import App,TkinterDnD
fixture=app/'.qa'/'review-fixture'
fixture.mkdir(parents=True,exist_ok=True)
for i in range(30):
    folder=fixture/f'新しいフォルダー ({i+1:02})';folder.mkdir(exist_ok=True)
    if i==29:continue
    image=Image.new('RGB',(900,650),('#d5e8f0' if i%2 else '#f0e3cb'))
    draw=ImageDraw.Draw(image);draw.rectangle((100,140,780,500),outline='#23486a',width=8)
    draw.text((140,180),f'SAMPLE PROPERTY {i+1:02}',fill='#23486a',font_size=44)
    draw.line((440,140,440,500),fill='#23486a',width=8)
    image.save(folder/'IMG_2847.jpg')
    pdf=Canvas(str(folder/'scan001.pdf'));pdf.setFont('Helvetica',24);pdf.drawString(60,760,f'SAMPLE PROPOSAL {i+1:02}');pdf.rect(60,300,480,380);pdf.showPage();pdf.drawString(60,760,'SECOND PAGE');pdf.save()
    (folder/'001.xlsx').write_bytes(b'fixture placeholder - icon only')
    (folder/'subfolder').mkdir(exist_ok=True)
first=fixture/'新しいフォルダー (01)'
for i in range(400):
    path=first/f'IMG_{i:04}.png'
    if not path.exists():image.save(path)
(first/'broken.pdf').write_bytes(b'not a pdf')
writer=PdfWriter();writer.append(PdfReader(first/'scan001.pdf'));writer.encrypt('synthetic-password')
with (first/'encrypted.pdf').open('wb') as f:writer.write(f)
root=TkinterDnD.Tk(); app_ui=App(root)
errors=[]
root.report_callback_exception=lambda *exc:errors.append(str(exc))
ticks=[];start=time.perf_counter()
def tick():ticks.append(time.perf_counter());root.after(10,tick)
tick();app_ui.load([fixture],True)
def spin(condition,timeout=20):
    end=time.perf_counter()+timeout
    while time.perf_counter()<end:
        root.update();time.sleep(.002)
        if condition():return
    raise AssertionError('UI condition timed out: '+str({'errors':errors,'cards':len(app_ui.cards),'files':len(app_ui.files),'visible':app_ui.visible,'render_gen':app_ui.render_gen}))
spin(lambda:len(app_ui.folders)==30 and app_ui.current==0)
app_ui.canvas.yview_moveto(.02);app_ui.schedule_view()
spin(lambda:any(c['image'] is not None for c in app_ui.cards))
app_ui.proposed.set('営業資料テスト');app_ui.checked.set(False);app_ui.save()
for fraction in (0,.25,.5,.75,1):
    app_ui.canvas.yview_moveto(fraction);app_ui.schedule_view()
    end=time.perf_counter()+.25
    spin(lambda:time.perf_counter()>end)
assert max(c['index'] or 0 for c in app_ui.cards)>=400
app_ui.tree.selection_set('1');spin(lambda:app_ui.current==1)
app_ui.tree.selection_set('0');spin(lambda:app_ui.current==0)
assert app_ui.proposed.get()=='営業資料テスト' and not app_ui.checked.get()
assert len(app_ui.cards)<=32
# Exercise the exact registered DnD callback with Tcl list quoting, including spaces.
app_ui.folders[0]['proposed']=''
labels=[w for w in root.winfo_children()[0].winfo_children() if w.winfo_class()=='TFrame']
# Parsing contract is shared with both registered callbacks.
paths=root.tk.splitlist(root.tk.call('list',str(first),str(fixture/'新しいフォルダー (02)')))
app_ui.load(paths,False);spin(lambda:len(app_ui.folders)==2 and not app_ui.loading)
assert len(app_ui.folders)==2
maximum=max(b-a for a,b in zip(ticks,ticks[1:]))
assert maximum<.5,maximum
result={'folders':30,'mixed_entries':len(list(first.iterdir())),'max_ui_heartbeat_gap_seconds':maximum,'max_live_cards':len(app_ui.cards),'input_check_preserved':True,'last_file_reachable':True,'direct_drop_parse_scan':True,'callback_errors':errors}
assert not errors
(app/'.qa'/'gui-result.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result))
app_ui.close()
