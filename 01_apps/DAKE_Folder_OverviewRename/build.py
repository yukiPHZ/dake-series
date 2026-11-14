"""Clean isolated candidate build; no deletion of previous candidates."""
from pathlib import Path
import datetime
import importlib.metadata as metadata
import subprocess
import sys
import shutil
import os

app=Path(__file__).resolve().parent
stamp=datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
build=app/'build'/stamp
build.mkdir(parents=True)
os.environ['PYINSTALLER_CONFIG_DIR']=str(build/'cache')
subprocess.run([sys.executable,str(app/'../../tools/generate_version_info.py'),'--app-dir',str(app),'--out',str(build/'version_info.txt')],check=True)
licenses=app/'third_party_licenses'
licenses.mkdir(exist_ok=True)
notices=[]
for name in ('pypdfium2','Pillow','tkinterdnd2'):
    dist=metadata.distribution(name)
    notices.append(f'{name} {dist.version}')
    for file in dist.files or []:
        if any(word in str(file).lower() for word in ('license','copying','copyright')):
            source=Path(dist.locate_file(file))
            if source.is_file():
                target=licenses/name/str(file).replace('..','_')
                target.parent.mkdir(parents=True,exist_ok=True)
                shutil.copyfile(source,target)
(app/'THIRD_PARTY_NOTICES.txt').write_text('DakeFolder_OverviewRename v1.0.0\nRuntime dependencies and bundled license texts:\n'+'\n'.join(notices)+'\nSee third_party_licenses. Python and Tcl/Tk are bundled by PyInstaller.\n',encoding='utf-8')
subprocess.run([sys.executable,'-m','PyInstaller','--onefile','--noconsole','--clean','--noconfirm','--name','DakeFolder_OverviewRename','--icon',str(app/'../../02_assets/dake_icon.ico'),'--version-file',str(build/'version_info.txt'),'--add-data',str(app/'../../02_assets/dake_icon.ico')+';.','--add-data',str(licenses)+';third_party_licenses','--add-data',str(app/'THIRD_PARTY_NOTICES.txt')+';.','--collect-all','tkinterdnd2','--workpath',str(build/'work'),'--specpath',str(build),'--distpath',str(app/'dist'/stamp),str(app/'main.py')],check=True,cwd=app)
print('CANDIDATE='+str(app/'dist'/stamp/'DakeFolder_OverviewRename.exe'))
