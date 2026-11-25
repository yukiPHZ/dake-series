import hashlib
import os
import sys
import time
from pathlib import Path
from unittest.mock import Mock
import pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from rename_core import Snapshot,Change,Journal,SafetyError,validate,validate_name,execute
from workers import scan,Renderer,LatestWorker
from PIL import Image

def folder(root,name='新しいフォルダー'):
    p=root/name;p.mkdir();(p/'scan001.txt').write_bytes(b'unchanged');return p

def hashes(root):
    return sorted(hashlib.sha256(p.read_bytes()).hexdigest() for p in root.rglob('*') if p.is_file())

@pytest.mark.parametrize('name',['','CON','nul.txt','COM1','LPT9','COM¹','.','..','a/b','a\\b','a:','a?','a.','a ','x\x00','😀'*128])
def test_invalid(name):
    with pytest.raises(SafetyError):validate_name(name)

@pytest.mark.parametrize('name',['日本語の長い名前'*20,'営業資料 2026','提案.確定','😀'*100])
def test_valid(name):validate_name(name)

@pytest.mark.parametrize('count',[0,1,10,30,100])
def test_inventory(tmp_path,count):
    for i in range(count):folder(tmp_path,str(i))
    rows,errors=scan(([tmp_path],True),lambda:False)
    assert len(rows)==count and not errors

def test_empty_and_nested(tmp_path):
    p=tmp_path/'empty';p.mkdir();(p/'sub').mkdir();(p/'sub'/'hidden.txt').write_text('not scanned')
    rows,errors=scan(([tmp_path],True),lambda:False)
    assert rows[0]['subdirs']==1 and len(rows[0]['files'])==1

def test_hash_batch_undo(tmp_path):
    root=tmp_path/'data';root.mkdir(); paths=[folder(root,str(i)) for i in range(30)]
    before=hashes(root)
    changes=[Change(Snapshot.capture(p),'営業資料'+p.name) for p in paths]
    results,undo=execute(changes,Journal(tmp_path/'logs'))
    assert len(undo)==30 and all(r[2]=='success' for r in results)
    assert hashes(root)==before
    results,redo=execute(undo,Journal(tmp_path/'logs'))
    assert len(redo)==30 and hashes(root)==before and all(p.exists() for p in paths)

def test_collision_external_and_nested(tmp_path):
    a=folder(tmp_path,'a'); b=folder(tmp_path,'b');snap=Snapshot.capture(a)
    with pytest.raises(SafetyError):validate([Change(snap,'b')])
    with pytest.raises(SafetyError):validate([Change(snap,'A')])
    with pytest.raises(SafetyError):validate([Change(snap,'c'),Change(Snapshot.capture(b),'C')])
    (a/'scan001.txt').write_bytes(b'changed')
    with pytest.raises(SafetyError):validate([Change(snap,'c')])

def test_replace_directory(tmp_path):
    a=folder(tmp_path);snap=Snapshot.capture(a);a.rename(tmp_path/'external');a.mkdir()
    with pytest.raises(SafetyError):validate([Change(snap,'new')])

def test_partial_and_undo_collision(tmp_path):
    data=tmp_path/'data';data.mkdir();a=folder(data,'a');b=folder(data,'b')
    calls=[]
    def move(src,dst):
        calls.append(src)
        if len(calls)==2:raise PermissionError('simulated sharing violation')
        os.rename(src,dst)
    results,undo=execute([Change(Snapshot.capture(p),p.name+'new') for p in (a,b)],Journal(tmp_path/'logs'),move)
    assert len(undo)==1 and b.exists() and not a.exists() and results[-1][2].startswith('failed')
    a.mkdir()
    with pytest.raises(SafetyError):validate(undo)

def test_undo_external(tmp_path):
    a=folder(tmp_path);results,undo=execute([Change(Snapshot.capture(a),'new')],Journal(tmp_path/'logs'))
    (tmp_path/'new'/'new.txt').write_text('external')
    with pytest.raises(SafetyError):validate(undo)

def test_journal_failure_before_move(tmp_path):
    a=folder(tmp_path);journal=Mock();journal.write.side_effect=OSError('disk full')
    results,undo=execute([Change(Snapshot.capture(a),'new')],journal)
    assert a.exists() and not undo

def test_journal_failure_after_move(tmp_path):
    a=folder(tmp_path);journal=Mock();journal.write.side_effect=[None,OSError('disk full'),None]
    results,undo=execute([Change(Snapshot.capture(a),'new')],journal)
    assert len(undo)==1 and results[0][2].startswith('renamed_log_failed')

def test_windows_lock(tmp_path):
    from win_rename import locked_directory
    a=folder(tmp_path)
    with locked_directory(a):
        results,undo=execute([Change(Snapshot.capture(a),'new')],Journal(tmp_path/'logs'))
    assert not undo and a.exists()

def test_junction(tmp_path):
    import subprocess
    target=tmp_path/'target';target.mkdir();link=tmp_path/'link'
    result=subprocess.run(['cmd','/c','mklink','/J',str(link),str(target)],capture_output=True)
    assert result.returncode==0
    try:
        with pytest.raises(SafetyError):Snapshot.capture(link)
        rows,errors=scan(([link],False),lambda:False)
        assert not rows and errors
    finally:os.rmdir(link)

def test_mixed_many_and_corrupt(tmp_path):
    a=tmp_path/'data';a.mkdir()
    for i in range(310):(a/f'{i:04}.txt').write_text('sample')
    Image.new('RGB',(800,600),'red').save(a/'IMG_001.jpg')
    (a/'scan001.pdf').write_bytes(b'corrupt')
    rows,errors=scan(([a],False),lambda:False)
    assert len(rows[0]['files'])==312 and not errors
    renderer=Renderer()
    items=rows[0]['files']
    result=list(renderer(list(enumerate(items)),lambda:False))[-1]
    assert len(result)==312 and len(renderer.cache)<=64
    assert renderer.render(next(f for f in items if f['name']=='scan001.pdf'))[2]=='preview_failed'
    assert renderer.render(next(f for f in items if f['name']=='IMG_001.jpg'))[0].size==(293,220)

def test_latest_worker():
    worker=LatestWorker(lambda value,cancelled:(time.sleep(.03),value)[1])
    for i in range(50):generation=worker.submit(i)
    time.sleep(.15)
    gen,result,error=worker.output.get(timeout=2)
    assert gen==generation and result==49 and error is None
    worker.close();worker.thread.join(1);assert not worker.thread.is_alive()


def test_pdf_pages_encrypted_and_release_handles(tmp_path):
    from reportlab.pdfgen.canvas import Canvas
    from pypdf import PdfWriter,PdfReader
    a=tmp_path/'pdfs';a.mkdir()
    pdf=Canvas(str(a/'scan.pdf'));pdf.drawString(50,700,'SYNTHETIC');pdf.showPage();pdf.drawString(50,700,'PAGE 2');pdf.save()
    writer=PdfWriter();writer.append(PdfReader(a/'scan.pdf'));writer.encrypt('test-only')
    with (a/'encrypted.pdf').open('wb') as stream:writer.write(stream)
    rows,_=scan(([a],False),lambda:False)
    renderer=Renderer();rendered=list(renderer(list(enumerate(rows[0]['files'])),lambda:False))[-1]
    by_name={rows[0]['files'][index]['name']:value for index,value in rendered}
    assert by_name['scan.pdf'][1]==2 and by_name['scan.pdf'][0] is not None
    assert by_name['encrypted.pdf'][2]=='preview_failed'
    result,undo=execute([Change(rows[0]['snapshot'],'renamed')],Journal(tmp_path/'logs'))
    assert len(undo)==1


def test_nested_selection_and_long_actual_name(tmp_path):
    a=folder(tmp_path,'a'); child=folder(a,'child')
    with pytest.raises(SafetyError):validate([Change(Snapshot.capture(a),'new'),Change(Snapshot.capture(child),'subnew')])
    name='日本語営業資料'*20
    result,undo=execute([Change(Snapshot.capture(a),name)],Journal(tmp_path/'logs'))
    assert len(undo)==1 and (tmp_path/name).exists()


def test_native_no_replace_race(tmp_path):
    from win_rename import rename_checked
    a=folder(tmp_path);change=Change(Snapshot.capture(a),'destination')
    def check(changes):
        validate(changes)
        change.destination.mkdir()
    with pytest.raises(OSError):rename_checked(change,check)
    assert a.exists() and change.destination.exists()


def test_file_identity_external_change_same_metadata(tmp_path):
    a=folder(tmp_path);snap=Snapshot.capture(a)
    file=a/'scan001.txt';metadata=file.stat();file.rename(a/'old')
    file.write_bytes(b'unchanged');os.utime(file,ns=(metadata.st_atime_ns,metadata.st_mtime_ns))
    (a/'old').rename(tmp_path/'preserved')
    with pytest.raises(SafetyError):snap.verify()
