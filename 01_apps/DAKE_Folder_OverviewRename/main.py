import os
import sys
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import ImageTk
from workers import LatestWorker, Renderer, scan
from rename_core import Change, Journal, validate, execute

APP_NAME = 'Dake名前フォルダ中身見ながら変更'
WINDOW_TITLE = APP_NAME + ' v1.0.0'
COPYRIGHT = '© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta'
UI_TEXT = {
 'title':APP_NAME,'description':'親フォルダを選ぶ → 中身を見る → 名前を入力 → まとめて変更',
 'parent':'親フォルダを選ぶ','drop_parent':'親フォルダをここへドロップ',
 'drop_direct':'名前を変えたいフォルダをここへドロップ（複数可）',
 'apply':'変更内容を確認','undo':'直前の変更を元に戻す','refresh':'再読み込み',
 'name':'現在のフォルダ名','count':'ファイル / 子フォルダ','planned':'変更予定名',
 'check':'このフォルダを変更対象にする','next':'次のフォルダへ',
 'empty':'親フォルダを選択してください。中身を見ながら名前を付けられます。',
 'loading':'中身を読み込んでいます…','busy':'名前変更中です…','ready':'{n} フォルダ / 読み込み不可 {errors} 件',
 'details':'内容を選択してください','summary':'{name}  —  ファイル {files} 件 / 子フォルダ {dirs} 件',
 'no_files':'空のフォルダです','no_selection':'左のフォルダを選ぶと中身を表示します','pending':'プレビュー準備中',
 'folder':'▣ 子フォルダ（走査しません）','link':'↗ リンク（読み込みません）',
 'other':'▤ {kind} ファイル','pdf':'PDF','image':'画像','preview_failed':'プレビュー不可（破損・暗号化・非対応等）',
 'unreadable':'読み取り不可','size':'{size:,.0f} KB','pages':'{pages} ページ','unknown':'容量不明',
 'error':'処理を確認してください','confirm':'変更前の確認','execute':'この内容で変更する','cancel':'閉じる',
 'confirm_hint':'旧名 → 新名を確認してください。中のファイルは変更しません。',
 'result':'変更結果','result_summary':'処理済み {done} / 対象 {total}\n未処理分は変更していません。\n履歴: {log}',
 'no_changes':'変更予定名が入力された対象を選んでください。',
 'undo_empty':'元へ戻せる変更がありません。','undo_warning':'成功した変更を元に戻します。外部変更・衝突がある場合は停止します。',
 'replace':'読み込み直すと未実行の入力が消えます。続けますか？',
 'close_busy':'名前変更処理の完了を待ってから閉じてください。',
 'validation':'名前・衝突・外部変更を確認できませんでした。再読み込みして確認してください。\n{error}',
 'footer':'シンプルそれDAKEシリーズ　|　止まらない、迷わない、すぐ終わる。',
 'copyright':COPYRIGHT,'success':'変更済み','failed':'失敗・要確認','skipped':'未実行',
 'checkbox_on':'✓ ','checkbox_off':'— ', 'files_count':'{files} / {dirs}', 'sample':'内容の手がかり', 'preview_names':'{names}',
 'footer_link_1':'戸建買取査定','footer_link_2':'Instagram',
 'footer_url_1':'https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74',
 'footer_url_2':'https://instagram.com/kikuta.shimarisu_fudosan',
}

class App:
    def __init__(self, root):
        self.root=root
        self.folders=[]; self.current=None; self.files=[]; self.undo=[]; self.request=None
        self.busy=False; self.loading=False; self.scan_gen=0; self.render_gen=0
        self.scanner=LatestWorker(scan); self.renderer=LatestWorker(Renderer())
        self.results=queue.Queue(); self.cards=[]; self.visible=None; self.scheduled=None
        self.root.title(WINDOW_TITLE); self.root.geometry('1240x820'); self.root.minsize(960,640)
        families=font.families()
        chosen=next((f for f in ('BIZ UDPGothic','Yu Gothic UI','Meiryo') if f in families),'TkDefaultFont')
        self.root.option_add('*Font',(chosen,10))
        style=ttk.Style(); style.theme_use('clam')
        style.configure('.',font=(chosen,10),background='#f4f5f7',foreground='#263445')
        style.configure('Treeview',rowheight=42,background='white',fieldbackground='white')
        style.configure('Accent.TButton',background='#245da8',foreground='white',padding=10)
        base=Path(getattr(sys,'_MEIPASS',Path(__file__).parent))
        for icon in (base/'dake_icon.ico',base/'../../02_assets/dake_icon.ico'):
            if icon.exists():
                try: root.iconbitmap(str(icon))
                except tk.TclError: pass
                break
        outer=ttk.Frame(root,padding=20); outer.pack(fill='both',expand=True)
        ttk.Label(outer,text=UI_TEXT['title'],font=(chosen,17,'bold')).pack(anchor='w')
        ttk.Label(outer,text=UI_TEXT['description']).pack(anchor='w',pady=(8,12))
        bar=ttk.Frame(outer); bar.pack(fill='x')
        self.parent_button=ttk.Button(bar,text=UI_TEXT['parent'],command=self.choose); self.parent_button.pack(side='left')
        self.refresh_button=ttk.Button(bar,text=UI_TEXT['refresh'],command=self.refresh); self.refresh_button.pack(side='left',padx=8)
        self.apply_button=ttk.Button(bar,text=UI_TEXT['apply'],style='Accent.TButton',command=self.prepare); self.apply_button.pack(side='right')
        self.undo_button=ttk.Button(bar,text=UI_TEXT['undo'],command=self.prepare_undo); self.undo_button.pack(side='right',padx=8)
        drops=ttk.Frame(outer); drops.pack(fill='x',pady=10)
        for key,parent in (('drop_parent',True),('drop_direct',False)):
            label=tk.Label(drops,text=UI_TEXT[key],background='#e7edf5',foreground='#31567c',padx=12,pady=12)
            label.pack(side='left',fill='x',expand=True,padx=3)
            label.drop_target_register(DND_FILES)
            label.dnd_bind('<<Drop>>',lambda e,p=parent:self.load(self.root.tk.splitlist(e.data),p))
        pane=ttk.Panedwindow(outer,orient='horizontal'); pane.pack(fill='both',expand=True)
        left=ttk.Frame(pane); right=ttk.Frame(pane); pane.add(left,weight=2); pane.add(right,weight=3)
        self.tree=ttk.Treeview(left,columns=('count','planned','sample'),show='tree headings',selectmode='browse',height=5)
        self.tree.heading('#0',text=UI_TEXT['name']); self.tree.column('#0',width=150,minwidth=100)
        for column,key,width in (('count','count',80),('planned','planned',130),('sample','sample',120)):
            self.tree.heading(column,text=UI_TEXT[key]); self.tree.column(column,width=width,minwidth=80)
        tree_scroll=ttk.Scrollbar(left,command=self.tree.yview); self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side='right',fill='y'); self.tree.pack(fill='both',expand=True)
        self.tree.bind('<<TreeviewSelect>>',self.select)
        editor=ttk.Frame(left,padding=(0,12)); editor.pack(fill='x')
        self.checked=tk.BooleanVar(value=True); self.proposed=tk.StringVar()
        ttk.Checkbutton(editor,text=UI_TEXT['check'],variable=self.checked,command=self.save).pack(anchor='w')
        ttk.Label(editor,text=UI_TEXT['planned']).pack(anchor='w',pady=(10,4))
        self.entry=ttk.Entry(editor,textvariable=self.proposed); self.entry.pack(fill='x')
        self.proposed.trace_add('write',lambda *_:self.save())
        ttk.Button(editor,text=UI_TEXT['next'],command=self.next).pack(anchor='e',pady=6)
        self.detail=ttk.Label(right,text=UI_TEXT['details'],wraplength=560); self.detail.pack(anchor='w',padx=12,pady=8)
        viewport=ttk.Frame(right); viewport.pack(fill='both',expand=True)
        self.canvas=tk.Canvas(viewport,background='#eef0f3',highlightthickness=0)
        scrollbar=ttk.Scrollbar(viewport,command=self.scroll); scrollbar.pack(side='right',fill='y')
        self.canvas.pack(fill='both',expand=True,padx=(10,0)); self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>',lambda e:self.schedule_view())
        self.canvas.bind('<MouseWheel>',self.wheel)
        self.canvas_empty=self.canvas.create_text(25,40,text=UI_TEXT['no_files'],anchor='nw',fill='#64748b')
        self.status=tk.StringVar(value=UI_TEXT['empty']); ttk.Label(outer,textvariable=self.status).pack(anchor='w',pady=(10,6))
        ttk.Label(outer,text=UI_TEXT['footer']).pack(anchor='center')
        links=ttk.Frame(outer);links.pack()
        import webbrowser
        for key in (1,2):
            ttk.Button(links,text=UI_TEXT[f'footer_link_{key}'],command=lambda k=key:webbrowser.open(UI_TEXT[f'footer_url_{k}'])).pack(side='left',padx=4)
        ttk.Label(outer,text=UI_TEXT['copyright'],font=(chosen,8)).pack(anchor='center',pady=(4,0))
        root.protocol('WM_DELETE_WINDOW',self.close)
        self.timer=root.after(40,self.poll)
    def save(self):
        if self.current is None or self.busy: return
        row=self.folders[self.current]; row['proposed']=self.proposed.get(); row['checked']=self.checked.get()
        self.update_row(self.current)
    def update_row(self,index):
        row=self.folders[index]
        self.tree.item(str(index),text=UI_TEXT['checkbox_on' if row['checked'] else 'checkbox_off']+row['snapshot'].path.name,values=(UI_TEXT['files_count'].format(files=sum(not f['directory'] for f in row['files']),dirs=row['subdirs']),row['proposed'],UI_TEXT['preview_names'].format(names=' / '.join(f['name'] for f in row['files'][:2]))))
    def choose(self):
        if self.busy:return
        path=filedialog.askdirectory()
        if path:self.load([path],True)
    def refresh(self):
        if self.request:self.load(*self.request)
    def load(self,paths,parent):
        if self.busy:return
        if any(r['proposed'] for r in self.folders) and not messagebox.askyesno(UI_TEXT['confirm'],UI_TEXT['replace']):return
        self.request=(tuple(paths),parent); self.loading=True
        self.scan_gen=self.scanner.submit(self.request)
        self.render_gen=self.renderer.submit([])
        self.current=None; self.files=[]; self.visible=None; self.schedule_view()
        self.status.set(UI_TEXT['loading']); self.apply_button.state(['disabled'])
    def select(self,event=None):
        if self.busy or self.loading:return
        selected=self.tree.selection()
        if not selected:return
        index=int(selected[0]); self.current=None
        row=self.folders[index]; self.proposed.set(row['proposed']); self.checked.set(row['checked']); self.current=index
        self.files=row['files']; self.visible=None; self.canvas.yview_moveto(0)
        self.detail.configure(text=UI_TEXT['summary'].format(name=row['snapshot'].path.name,files=sum(not f['directory'] for f in self.files),dirs=row['subdirs']))
        self.render_gen=self.renderer.submit([]); self.schedule_view()
    def next(self):
        if self.folders and not self.busy:
            index=min((self.current or 0)+1,len(self.folders)-1)
            self.tree.selection_set(str(index)); self.tree.see(str(index)); self.entry.focus_set()
    def scroll(self,*args):self.canvas.yview(*args); self.schedule_view()
    def wheel(self,event):
        self.canvas.yview_scroll(-int(event.delta/120),'units'); self.schedule_view(); return 'break'
    def schedule_view(self):
        if self.scheduled is None:self.scheduled=self.root.after(15,self.draw)
    def draw(self):
        self.scheduled=None
        width=max(300,self.canvas.winfo_width()); columns=max(1,width//310); height=300
        self.canvas.configure(scrollregion=(0,0,width,max(1,(len(self.files)+columns-1)//columns)*height))
        start=max(0,int(self.canvas.canvasy(0)//height)*columns)
        count=min(32,((self.canvas.winfo_height()//height)+2)*columns)
        indices=tuple(range(start,min(len(self.files),start+count)))
        if self.visible==(indices,columns,width):return
        self.visible=(indices,columns,width)
        self.canvas.itemconfigure(self.canvas_empty,state='hidden' if self.files else 'normal',text=UI_TEXT['no_files' if self.current is not None else 'no_selection'])
        while len(self.cards)<len(indices):
            frame=tk.Frame(self.canvas,bg='white',bd=1,relief='solid')
            picture=tk.Label(frame,bg='white',fg='#61758b'); picture.place(x=6,y=6,relwidth=1,width=-12,height=220)
            name=tk.Label(frame,bg='white',anchor='w',justify='left'); name.place(x=10,y=228,relwidth=1,width=-20,height=36)
            meta=tk.Label(frame,bg='white',anchor='w',fg='#64748b'); meta.place(x=10,y=268,relwidth=1,width=-20,height=22)
            for widget in (frame,picture,name,meta):widget.bind('<MouseWheel>',self.wheel)
            window=self.canvas.create_window(0,0,window=frame,anchor='nw')
            self.cards.append({'window':window,'picture':picture,'name':name,'meta':meta,'image':None,'index':None})
        for slot,card in enumerate(self.cards):
            if slot>=len(indices):self.canvas.itemconfigure(card['window'],state='hidden'); card['image']=None;card['index']=None;continue
            index=indices[slot]; item=self.files[index]; card['index']=index;card['image']=None
            card['picture'].configure(image='',text=UI_TEXT['pending'])
            card['name'].configure(text=item['name'],wraplength=width//columns-30)
            card['meta'].configure(text=self.size_text(item))
            self.canvas.itemconfigure(card['window'],state='normal',width=width//columns-12,height=294)
            self.canvas.coords(card['window'],(index%columns)*(width//columns)+6,(index//columns)*height+3)
        self.render_gen=self.renderer.submit([(i,self.files[i]) for i in indices])
    def size_text(self,item):
        return UI_TEXT['unknown'] if item['size'] is None else UI_TEXT['size'].format(size=item['size']/1024)
    def poll(self):
        try:
            gen,result,error=self.scanner.output.get_nowait()
            if gen==self.scan_gen:
                self.loading=False; self.apply_button.state(['!disabled'])
                if error:messagebox.showerror(UI_TEXT['error'],UI_TEXT['validation'].format(error=error))
                else:
                    self.current=None; self.folders,failures=result
                    self.tree.delete(*self.tree.get_children())
                    for index,row in enumerate(self.folders):self.tree.insert('', 'end',iid=str(index)); self.update_row(index)
                    self.status.set(UI_TEXT['ready'].format(n=len(self.folders),errors=len(failures)))
                    if self.folders:self.tree.selection_set('0')
                    else:self.detail.configure(text=UI_TEXT['no_files'])
                    if failures:self.report(UI_TEXT['error'], '\n'.join(p+'\n'+e for p,e in failures))
        except queue.Empty:pass
        try:
            gen,result,error=self.renderer.output.get_nowait()
            if gen==self.render_gen and result:
                lookup=dict(result)
                for card in self.cards:
                    if card['index'] not in lookup:continue
                    im,pages,kind=lookup[card['index']]
                    label=UI_TEXT[kind] if kind!='other' else UI_TEXT['other'].format(kind=self.files[card['index']]['path'].suffix.upper().lstrip('.') or 'FILE')
                    if im is not None:
                        photo=ImageTk.PhotoImage(im);card['image']=photo;card['picture'].configure(image=photo,text='')
                    else:card['picture'].configure(image='',text=label,wraplength=260)
                    meta=self.size_text(self.files[card['index']])
                    if pages is not None:meta+=' / '+UI_TEXT['pages'].format(pages=pages)
                    card['meta'].configure(text=meta)
        except queue.Empty:pass
        try:
            kind,payload=self.results.get_nowait()
            self.set_busy(False)
            if kind=='error':messagebox.showerror(UI_TEXT['error'],UI_TEXT['validation'].format(error=payload))
            elif kind=='plan':self.confirm(payload)
            else:
                results,undo,log,changes,is_undo=payload
                if is_undo:
                    completed={source for source,dest,state in results if state=='success' or state.startswith('renamed_log_failed')}
                    self.undo=[c for c in changes if str(c.snapshot.path) not in completed]
                else:self.undo=undo
                # Keep the model synchronized with only known successful filesystem changes.
                for inverse in undo:
                    for row in self.folders:
                        if row['snapshot'].ident==inverse.snapshot.ident:
                            row['snapshot']=inverse.snapshot; row['proposed']=''
                            for item in row['files']:item['path']=inverse.snapshot.path/item['name']
                for i in range(len(self.folders)):self.update_row(i)
                self.select()
                summary=UI_TEXT['result_summary'].format(done=len(undo),total=len(changes),log=log)
                self.status.set(summary.split('\n')[0]); self.report(UI_TEXT['result'],summary+'\n\n'+'\n'.join(a+' → '+b+'\n'+s for a,b,s in results))
        except queue.Empty:pass
        self.timer=self.root.after(40,self.poll)
    def set_busy(self,value):
        self.busy=value
        for widget in (self.parent_button,self.refresh_button,self.apply_button,self.undo_button,self.entry):widget.state(['disabled' if value else '!disabled'])
        if value:self.status.set(UI_TEXT['busy'])
    def prepare(self):
        if self.busy or self.loading:return
        changes=[Change(r['snapshot'],r['proposed']) for r in self.folders if r['checked'] and r['proposed'] and r['proposed']!=r['snapshot'].path.name]
        if not changes:messagebox.showinfo(UI_TEXT['confirm'],UI_TEXT['no_changes']);return
        self.check_plan(changes,False)
    def prepare_undo(self):
        if self.busy or self.loading:return
        if not self.undo:messagebox.showinfo(UI_TEXT['confirm'],UI_TEXT['undo_empty']);return
        self.check_plan(list(self.undo),True)
    def check_plan(self,changes,is_undo):
        self.set_busy(True)
        def job():
            try:validate(changes);self.results.put(('plan',(changes,is_undo)))
            except Exception as exc:self.results.put(('error',str(exc)))
        threading.Thread(target=job,daemon=True).start()
    def confirm(self,payload):
        changes,is_undo=payload
        dialog=self.report(UI_TEXT['confirm'],UI_TEXT['undo_warning' if is_undo else 'confirm_hint']+'\n\n'+'\n'.join(str(c.snapshot.path)+'\n  → '+c.name for c in changes))
        def run():
            dialog.destroy();self.set_busy(True)
            def job():
                try:
                    log=Journal(Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'DAKE_Folder_OverviewRename'/'history')
                    results,undo=execute(changes,log)
                    self.results.put(('done',(results,undo,log.path,changes,is_undo)))
                except Exception as exc:self.results.put(('error',str(exc)))
            threading.Thread(target=job,daemon=False).start()
        ttk.Button(dialog,text=UI_TEXT['execute'],style='Accent.TButton',command=run).pack(pady=8)
    def report(self,title,text):
        dialog=tk.Toplevel(self.root);dialog.title(title);dialog.geometry('760x520');dialog.transient(self.root);dialog.grab_set()
        body=ttk.Frame(dialog,padding=15);body.pack(fill='both',expand=True)
        box=tk.Text(body,wrap='word',background='white');scroll=ttk.Scrollbar(body,command=box.yview);box.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right',fill='y');box.pack(fill='both',expand=True);box.insert('1.0',text);box.configure(state='disabled')
        ttk.Button(dialog,text=UI_TEXT['cancel'],command=dialog.destroy).pack(pady=6)
        return dialog
    def close(self):
        if self.busy:messagebox.showinfo(UI_TEXT['confirm'],UI_TEXT['close_busy']);return
        self.scanner.close();self.renderer.close();self.root.after_cancel(self.timer)
        if self.scheduled:self.root.after_cancel(self.scheduled)
        self.root.destroy()

def main():
    root=TkinterDnD.Tk();app=App(root)
    if '--launch-check' in sys.argv:root.after(100,app.close)
    root.mainloop()
if __name__=='__main__':main()
