from pathlib import Path

p = Path(__file__).parent / '01_apps/DAKE_PDF_Compress/main.py'
s = p.read_text(encoding='utf-8')
s = s.replace('import os\n', 'import os\nimport multiprocessing as mp\n')
s = s.replace('from dataclasses import dataclass', 'from dataclasses import dataclass, field')
s = s.replace('    "status_idle":', '    "status_checking": "PDFを確認中...",\n    "status_analyzing": "PDFを解析中",\n    "status_images": "画像を最適化中",\n    "status_optimizing": "PDFを整えています",\n    "status_verifying": "仕上がりを確認中",\n    "status_saving": "保存中",\n    "status_closing": "処理を終了しています",\n    "result_sizes": "{before} → {after}",\n    "result_reduction": "{rate:.1f}%軽くなりました",\n    "status_idle":')
s = s.replace('"message_complete": "PDFの圧縮が完了しました。",', '"message_complete": "{rate:.1f}%軽くなりました",')
s = s.replace('"message_complete_detail": "保存先フォルダを開きます。",', '"message_complete_detail": "圧縮したPDFを同じフォルダに保存しました。",')
s = s.replace('このPDFはあまり圧縮できませんでした。すでに圧縮済み、またはPDF構造上、削減幅が小さい可能性があります。', 'このPDFは、すでにかなり軽いようです。')
s = s.replace('このPDFは圧縮効果がありませんでした。すでに圧縮済み、またはPDF構造上、削減幅が小さい可能性があります。', 'このPDFは、すでにかなり軽いようです。元PDFをご利用ください。')
s = s.replace('PDF処理に必要なライブラリが見つかりません。requirements.txt をインストールしてください。', 'PDF処理の準備ができません。アプリを入れ直してください。')
s = '\n'.join(line for line in s.split('\n') if not any(x in line for x in ['"message_fallback_used":', '"detail_suffix":', '"status_processing_base":', '"status_processing_dots":']))
s = s.replace('GHOSTSCRIPT_LIGHTER_PDF_SETTINGS = "/screen"\n', '')
s = s.replace('STATUS_PHRASE_DELAY_SECONDS = 1.6', 'STATUS_PHRASE_DELAY_SECONDS = 8.0')
s = s.replace('  Uses Ghostscript first when available, then built-in fallback.', '  Automatically selects a verified compression result.')
s = s.replace('    ghostscript_path: str | None = None\n', '    ghostscript_path: str | None = None\n    diagnostics: dict = field(default_factory=dict)\n')
s = s.replace('            fitz = fitz_module', '            fitz_module.TOOLS.mupdf_display_errors(False)\n            fitz_module.TOOLS.mupdf_display_warnings(False)\n            fitz = fitz_module')
s = s.replace('        print(f"[DakePDF_Compress] {message}")', '        write_app_debug_log(message)')
s = s.replace('        if doc.page_count < 1:', '        if not doc.is_pdf or doc.page_count < 1:')
start = s.index('def rewrite_images_if_supported(')
end = s.index('def cli_write_error(', start)
s = s[:start] + '''def compress_pdf(source_path: Path, phase=lambda key: None) -> PdfResult:
    from adaptive import compress, library
    try:
        library()
    except (ImportError, RuntimeError) as exc:
        raise CompressError("error_dependency_missing", str(exc)) from exc
    validate_pdf(source_path)
    gs = find_ghostscript()
    try:
        selected, profile, candidates = compress(source_path, gs, phase)
        diagnostics = {"profile": profile, "candidates": candidates,
                       "ghostscript": str(gs) if gs else None}
        write_app_debug_log("adaptive compression", context=diagnostics)
        if selected is None:
            key = "error_no_reduction" if any(c.get("checks") == "passed" for c in candidates) else "error_unknown"
            raise CompressError(key)
        rate = (1 - selected["size"] / profile["bytes"]) * 100
        return PdfResult(Path(selected["path"]), profile["bytes"], selected["size"],
                         rate, rate < LOW_REDUCTION_THRESHOLD, selected["name"],
                         ghostscript_path=str(gs) if gs else None, diagnostics=diagnostics)
    except CompressError:
        raise
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except Exception as exc:
        write_app_debug_log("compression failed", exc=exc)
        raise CompressError("error_unknown", str(exc)) from exc


def pdf_worker(connection):
    """Exactly one GUI worker; all PDF access stays outside the Tk process."""
    try:
        while True:
            command, generation, path = connection.recv()
            if command == "stop":
                return
            try:
                if command == "check":
                    validate_pdf(path)
                    payload = (path, path.stat().st_size, unique_output_path(path))
                    connection.send(("checked", (generation, payload)))
                else:
                    result = compress_pdf(path, lambda key: connection.send(("phase", key)))
                    connection.send(("success", result))
            except Exception as exc:
                write_app_debug_log("PDF worker failed", exc=exc)
                key = exc.message_key if isinstance(exc, CompressError) else "error_unknown"
                connection.send(("check_error" if command == "check" else "error", (generation, key)))
    except (EOFError, BrokenPipeError, OSError):
        pass
    finally:
        connection.close()


''' + s[end:]
s=s.replace('        self.is_processing = False\n        self.event_queue', '''        self.is_processing = False
        self.is_checking = False
        self.generation = 0
        self.pending_check = None
        self.worker_process = None
        self.worker_connection = None
        self.worker_busy = False
        self.worker_started = 0.0
        self.closing = False
        self.phase_key = "status_processing"
        self.event_queue''')
s=s.replace('        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)', '        self.poll_after_id = self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)')
s=s.replace('        self.root.bind("<Configure>", self.handle_root_configure)', '''        self.root.bind("<Configure>", self.handle_root_configure)
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)
        # Developer-only readiness evidence: no effect unless explicitly enabled.
        self.root.after_idle(self.record_ready)''')
s=s.replace('        self.add_info_row(info, 2, 0, UI_TEXT["label_compressed_size"], self.compressed_size_var)', '        self.add_info_row(info, 2, 0, UI_TEXT["label_compressed_size"], self.compressed_size_var)')
s=s.replace('font=(self.font_family, 11, "bold"),\n            wraplength=330,', 'font=(self.font_family, 14 if value_var is self.reduction_rate_var else 11, "bold"),\n            wraplength=290,')
s=s.replace('        self.update_action_state()\n\n    def add_info_row', '''        for widget in self.drop_area.winfo_children():
            widget.bind("<Button-1>", self.select_pdf_dialog)
        self.update_action_state()

    def add_info_row''')
s=s.replace('        files = [path for path in paths if path.is_file()]\n        if len(files) != 1:', '        if len(paths) != 1:')
s=s.replace('        self.load_pdf(files[0])', '        self.load_pdf(paths[0])')
start=s.index('    def load_pdf(')
end=s.index('    def clear_selection(',start)
s=s[:start]+'''    def load_pdf(self, path: Path) -> None:
        if self.is_processing or self.closing:
            return
        self.generation += 1
        self.selected_pdf = None
        self.is_checking = True
        self.file_name_var.set(truncate_middle(path.name, 58))
        self.original_size_var.set(UI_TEXT["value_not_yet"])
        self.save_name_var.set(UI_TEXT["value_not_yet"])
        self.save_folder_var.set(truncate_middle(str(path.parent), 70))
        self.compressed_size_var.set(UI_TEXT["value_not_yet"])
        self.reduction_rate_var.set(UI_TEXT["value_not_yet"])
        self.notice_var.set("")
        self.drop_title_var.set(UI_TEXT["status_checking"])
        self.drop_subtitle_var.set(truncate_middle(path.name, 58))
        self.set_status("status_checking", "processing")
        self.progress.start(10)
        self.update_action_state()
        self.pending_check = ("check", self.generation, path)
        # Paint first, then create / dispatch to the worker.
        self.root.after_idle(self.dispatch_pending)

    def ensure_worker(self) -> None:
        if self.worker_process is not None and self.worker_process.is_alive():
            return
        ctx = mp.get_context("spawn")
        parent, child = ctx.Pipe()
        self.worker_connection = parent
        self.worker_process = ctx.Process(target=pdf_worker, args=(child,))
        self.worker_process.start()
        child.close()

    def dispatch_pending(self) -> None:
        if self.closing or self.worker_busy or self.pending_check is None:
            return
        try:
            self.ensure_worker()
            self.worker_connection.send(self.pending_check)
            self.pending_check = None
            self.worker_busy = True
            self.worker_started = time.monotonic()
        except Exception as exc:
            self.is_checking = False
            self.worker_busy = False
            self.show_error("error_unknown", str(exc))
            self.update_action_state()

    def apply_checked(self, payload) -> None:
        path, size, output_path = payload
        self.is_checking = False
        self.progress.stop()
        self.selected_pdf = path
        self.original_size_var.set(format_bytes(size))
        self.save_name_var.set(truncate_middle(output_path.name, 58))
        self.drop_title_var.set(UI_TEXT["drop_title_selected"])
        self.drop_subtitle_var.set(UI_TEXT["drop_subtitle_selected"])
        self.set_status("status_ready", "ready")
        self.update_action_state()

''' +s[end:]
s=s.replace('        self.selected_pdf = None\n        self.file_name_var', '''        self.generation += 1
        self.pending_check = None
        self.is_checking = False
        self.progress.stop()
        self.selected_pdf = None
        self.file_name_var''')
s=s.replace('        self.is_processing = True\n        self.notice_var', '        self.is_processing = True\n        self.phase_key = "status_analyzing"\n        self.notice_var')
start=s.index('        worker = threading.Thread(target=self.compress_worker')
end=s.index('    def handle_queue_event(',start)
s=s[:start]+'''        try:
            self.ensure_worker()
            self.worker_connection.send(("compress", self.generation, source_path))
            self.worker_busy = True
            self.worker_started = time.monotonic()
        except Exception as exc:
            self.handle_worker_error(CompressError("error_unknown", str(exc)))

    def poll_queue(self) -> None:
        if self.closing:
            return
        try:
            if self.worker_connection is not None:
                while self.worker_connection.poll():
                    event = self.worker_connection.recv()
                    if event[0] != "phase":
                        self.worker_busy = False
                    self.handle_queue_event(event)
            if self.worker_busy and self.worker_process is not None:
                limit = 160 if self.is_processing else 30
                if not self.worker_process.is_alive() or time.monotonic() - self.worker_started > limit:
                    raise RuntimeError("PDF worker stopped or exceeded time budget")
        except (EOFError, OSError, RuntimeError) as exc:
            self.stop_worker()
            self.is_checking = False
            self.handle_worker_error(CompressError("error_unknown", str(exc)))
        self.dispatch_pending()
        self.poll_after_id = self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def stop_worker(self) -> None:
        if self.worker_process is not None:
            if self.worker_process.is_alive():
                self.worker_process.terminate()
            self.worker_process.join(0.2)
            if not self.worker_process.is_alive():
                self.worker_process.close()
            self.worker_process = None
        if self.worker_connection is not None:
            self.worker_connection.close()
            self.worker_connection = None
        self.worker_busy = False

    def close_app(self) -> None:
        if self.is_processing:
            # Let the already-requested atomic save finish. Window stays responsive.
            self.set_status("status_closing", "processing")
            self.close_after_processing = True
            return
        self.closing = True
        self.generation += 1
        self.pending_check = None
        self.stop_status_animation()
        self.root.after_cancel(self.poll_after_id)
        self.stop_worker()
        self.root.destroy()

    def record_ready(self) -> None:
        path = os.environ.get("DAKE_STARTUP_READY_FILE")
        if path:
            Path(path).write_text(str(time.perf_counter_ns()), encoding="ascii")

''' +s[end:]
start=s.index('    def handle_queue_event(')
end=s.index('    def handle_success(',start)
s=s[:start]+'''    def handle_queue_event(self, event: tuple[str, Any]) -> None:
        event_type, payload = event
        if event_type == "checked":
            generation, value = payload
            if generation == self.generation:
                self.apply_checked(value)
        elif event_type == "check_error":
            generation, key = payload
            if generation == self.generation:
                self.is_checking = False
                self.progress.stop()
                self.drop_title_var.set(UI_TEXT["status_error"])
                self.show_error(key)
                self.update_action_state()
        elif event_type == "phase":
            self.phase_key = payload
            self.set_status(payload, "processing")
        elif event_type == "success":
            self.handle_success(payload)
        elif event_type == "error":
            self.handle_worker_error(CompressError(payload[1]))

''' +s[end:]
start=s.index('        self.reduction_rate_var.set(f"{result.reduction_rate:.1f}%")')
end=s.index('        self.update_action_state()\n        dialog(',start)
s=s[:start]+'''        summary = UI_TEXT["result_reduction"].format(rate=result.reduction_rate)
        self.reduction_rate_var.set(summary)
        self.drop_title_var.set(UI_TEXT["result_sizes"].format(
            before=format_bytes(result.original_size), after=format_bytes(result.compressed_size)))
        self.drop_subtitle_var.set(summary)
        self.save_name_var.set(truncate_middle(result.output_path.name, 58))
        self.save_folder_var.set(truncate_middle(str(result.output_path.parent), 70))
        self.notice_var.set(UI_TEXT["message_low_reduction"] if result.low_reduction else "")
        self.set_status("status_low_reduction" if result.low_reduction else "status_complete",
                        "warning" if result.low_reduction else "success")
        message = UI_TEXT["message_complete"].format(rate=result.reduction_rate)
        if result.low_reduction:
            message += "\\n\\n" + UI_TEXT["message_low_reduction"]
        message += "\\n\\n" + UI_TEXT["message_complete_detail"]
        dialog = messagebox.showwarning if result.low_reduction else messagebox.showinfo
''' +s[end:]
s=s.replace('        self.open_output_folder(result.output_path.parent)', '''        self.open_output_folder(result.output_path.parent)
        if getattr(self, "close_after_processing", False):
            self.close_app()''')
s=s.replace('        self.show_error(exc.message_key, exc.detail)\n\n    def open_output_folder', '''        self.show_error(exc.message_key, exc.detail)
        if getattr(self, "close_after_processing", False):
            self.close_app()

    def open_output_folder''')
s=s.replace('            message = f"{message}\\n\\n{UI_TEXT[\'detail_suffix\'].format(detail=detail)}"', '            write_app_debug_log("user-visible error", context={"key": message_key, "detail": detail})')
start=s.index('        sequence = list(UI_TEXT["status_processing_dots"])')
end=s.index('        self.status_animation_index += 1',start)
s=s[:start]+'''        self.status_var.set(UI_TEXT[self.phase_key] + "." * (self.status_animation_index % 3 + 1))
        if elapsed >= STATUS_PHRASE_DELAY_SECONDS:
            phrases = ["status_phrase_1", "status_phrase_2", "status_phrase_3"]
            self.notice_var.set(UI_TEXT[phrases[min(2, int((elapsed - STATUS_PHRASE_DELAY_SECONDS) / 4))]])
''' +s[end:]
s=s.replace('        has_pdf = self.selected_pdf is not None', '        has_pdf = self.selected_pdf is not None and not self.is_checking')
s=s.replace('state=tk.NORMAL if has_pdf else tk.DISABLED, cursor="hand2" if has_pdf else "arrow")', 'state=tk.NORMAL if has_pdf or self.is_checking else tk.DISABLED, cursor="hand2" if has_pdf or self.is_checking else "arrow")')
s=s.replace('    if get_fitz() is None:\n        raise CompressError("error_dependency_missing")\n\n\n\ndef create_launch_check_window', '    from adaptive import library\n    library()\n\n\n\ndef create_launch_check_window')
s=s.replace('    raise SystemExit(main())', '    mp.freeze_support()\n    raise SystemExit(main())')
# Dead imports and constants from the old sequential fallback implementation.
s=s.replace('import tempfile\n','').replace('import threading\n','').replace('import queue\n','')
s=s.replace('        self.event_queue: queue.Queue[tuple[str, Any]] = queue.Queue()\n','')
s=s.replace('from dake_quality_engine import atomic_replace,', 'from dake_quality_engine import')
start=s.index('def make_temp_pdf_path(');end=s.index('def find_ghostscript(',start);s=s[:start]+s[end:]
p.write_text(s,encoding='utf-8')
(p.parent/'requirements.txt').write_text('PyMuPDF==1.28.2\ntkinterdnd2==0.4.3\npyinstaller==6.20.0\n',encoding='utf-8')
