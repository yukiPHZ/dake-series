from pathlib import Path
p=Path(__file__).parent/'baseline/01_apps/DAKE_PDF_Compress/main.py'
s=p.read_text(encoding='utf-8')
point='        self.root.bind("<Configure>", self.handle_root_configure)'
s=s.replace(point,point+'\n        self.root.after_idle(self.record_ready)')
s=s.replace('    def setup_style(self) -> None:', '''    def record_ready(self):
        path = os.environ.get("DAKE_STARTUP_READY_FILE")
        if path:
            if not self.root.winfo_viewable():
                self.root.after(10, self.record_ready)
                return
            Path(path).write_text(str(time.perf_counter_ns()), encoding="ascii")
            if os.environ.get("DAKE_STARTUP_PROBE") == "1":
                self.root.after(50, self.root.destroy)

    def setup_style(self) -> None:''')
p.write_text(s,encoding='utf-8')
