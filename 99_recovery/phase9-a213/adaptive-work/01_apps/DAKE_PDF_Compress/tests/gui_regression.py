"""Integration test with real Tk and spawned PDF worker; dialogs recorded."""
from pathlib import Path
import json
import multiprocessing
import sys
import time
from types import SimpleNamespace
from unittest.mock import patch

APP=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(APP))
import main


def run(fixtures, output):
    root=main.make_root()
    app=main.DakePdfCompressApp(root)
    dialogs=[];folders=[];ticks=[];phases=[]
    main.messagebox.showinfo=lambda *x:dialogs.append(('info',x))
    main.messagebox.showwarning=lambda *x:dialogs.append(('warning',x))
    app.open_output_folder=lambda p:folders.append(str(p))
    original=app.set_status
    def status(key,state):
        phases.append(key);original(key,state)
    app.set_status=status
    def tick():
        ticks.append(time.monotonic())
        if not app.closing: root.after(20,tick)
    tick()
    def pump(until,timeout=180):
        start=time.monotonic()
        while not until():
            root.update();time.sleep(0.005)
            if time.monotonic()-start>timeout:raise AssertionError('UI timeout')
        root.update()
    root.update()
    assert (root.winfo_width(),root.winfo_height())==(860,740)
    assert root.minsize()==(760,720)
    assert main.resource_icon_path().is_file()
    assert app.footer.winfo_ismapped()
    start=time.monotonic()
    app.handle_drop(SimpleNamespace(data=root.tk.call('list',str(fixtures/'H_100_pages.pdf'))))
    immediate=time.monotonic()-start
    assert app.status_var.get()==main.UI_TEXT['status_checking']
    assert immediate<0.2,immediate
    # Latest request wins while an earlier validation is already dispatched.
    root.update()
    for name in ['A_text.pdf','D_color_scan.pdf','C_bitonal.pdf']:
        app.load_pdf(fixtures/name)
    pump(lambda:not app.is_checking)
    assert app.selected_pdf.name=='C_bitonal.pdf'
    worker_pid=app.worker_process.pid
    assert app.pending_check is None
    assert app.status_var.get()==main.UI_TEXT['status_ready']
    app.start_compression()
    pump(lambda:not app.is_processing)
    assert folders and dialogs[-1][0] in ('info','warning')
    assert '軽くなりました' in dialogs[-1][1][1]
    assert '→' in app.drop_title_var.get()
    assert 'Ghostscript' not in dialogs[-1][1][1]
    first_name=app.save_name_var.get()
    app.start_compression();pump(lambda:not app.is_processing)
    assert app.save_name_var.get()!=first_name
    assert app.worker_process.pid==worker_pid
    app.clear_selection();assert app.selected_pdf is None
    with patch.object(main.filedialog,'askopenfilename',return_value=str(fixtures/'A_text.pdf')):
        app.select_pdf_dialog()
    pump(lambda:not app.is_checking)
    assert app.selected_pdf.name=='A_text.pdf'
    # Clear invalidates in-flight validation; stale errors must not appear.
    app.load_pdf(fixtures/'missing.pdf');root.update();app.clear_selection()
    count=len(dialogs);pump(lambda:not app.worker_busy)
    assert len(dialogs)==count and app.selected_pdf is None
    app.load_pdf(fixtures/'missing.pdf');pump(lambda:not app.is_checking)
    assert dialogs[-1][0]=='warning' and app.selected_pdf is None
    root.geometry('760x720');root.update()
    assert app.footer.winfo_ismapped()
    assert app.footer.winfo_rooty()+app.footer.winfo_height()<=root.winfo_rooty()+root.winfo_height()
    assert app.footer_mode=='narrow'
    root.geometry('1000x740');root.update();assert app.footer_mode=='wide'
    gaps=[b-a for a,b in zip(ticks,ticks[1:])]
    assert max(gaps)<0.5,max(gaps)
    result={'passed':True,'drop_response_seconds':immediate,'max_20ms_tick_gap':max(gaps),
            'worker_count':1,'phases':list(dict.fromkeys(phases)), 'dialogs':dialogs,
            'folder_open_calls':folders,'dnd_registered':main.DND_ENABLED,
            'notes':'Real Tk, callback DnD and dialog selection; native Explorer drag tested separately.'}
    app.close_app()
    assert app.worker_process is None
    output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:result[k] for k in ['passed','drop_response_seconds','max_20ms_tick_gap','dnd_registered']}))


if __name__=='__main__':
    multiprocessing.freeze_support()
    run(Path(sys.argv[1]).resolve(),Path(sys.argv[2]).resolve())
