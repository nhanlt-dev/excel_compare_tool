import customtkinter as ctk
from tkinter import ttk, messagebox
import threading, queue, time
from core.excel_loader import load_excel_columns
from core.compare_engine import compare_tables
from core.result_writer import save_result_dialog
from ui.style import apply_style
from utils.config import load_config, save_config
from utils.helper import timestamp, open_containing_folder

import pandas as pd

# init from saved config if exists
cfg = load_config()
mode_default = cfg.get("appearance_mode","system")
theme_default = cfg.get("color_theme","blue")
apply_style(mode_default, theme_default)

class ExcelCompareApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Compare PRO 5.0")
        self.geometry("1200x760")

        # state
        self.fileA = None; self.fileB = None
        self.colsA = []; self.colsB = []
        self.pairs = []
        self.result_df = None

        # thread control
        self.queue = queue.Queue()
        self.worker = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()

        # UI
        self._build_ui()
        self.after(100, self._process_queue)

    def _build_ui(self):
        pad=10
        toolbar = ctk.CTkFrame(self)
        toolbar.pack(fill="x", padx=pad, pady=8)

        btnA = ctk.CTkButton(toolbar, text="📂 Chọn file A", command=self._choose_file_a, width=140)
        btnA.pack(side="left", padx=6)
        self.lblA = ctk.CTkLabel(toolbar, text="Chưa chọn file A")
        self.lblA.pack(side="left", padx=6)

        btnB = ctk.CTkButton(toolbar, text="📂 Chọn file B", command=self._choose_file_b, width=140)
        btnB.pack(side="left", padx=6)
        self.lblB = ctk.CTkLabel(toolbar, text="Chưa chọn file B")
        self.lblB.pack(side="left", padx=6)

        # theme selection
        self.appear_var = ctk.StringVar(value=mode_default)
        self.theme_var = ctk.StringVar(value=theme_default)
        ctk.CTkLabel(toolbar, text="Giao diện:").pack(side="right", padx=6)
        self.menu_mode = ctk.CTkOptionMenu(toolbar, values=["system","dark","light"], variable=self.appear_var, command=self._on_change_theme)
        self.menu_mode.pack(side="right", padx=6)
        ctk.CTkLabel(toolbar, text="Màu:").pack(side="right", padx=6)
        self.menu_theme = ctk.CTkOptionMenu(toolbar, values=["blue","green","dark-blue"], variable=self.theme_var, command=self._on_change_theme)
        self.menu_theme.pack(side="right", padx=6)

        # mapping area
        cfg = ctk.CTkFrame(self)
        cfg.pack(fill="x", padx=pad, pady=6)

        ctk.CTkLabel(cfg, text="Key A:").grid(row=0,column=0,padx=6,pady=6, sticky="w")
        self.optKeyA = ctk.CTkOptionMenu(cfg, values=[], width=180)
        self.optKeyA.grid(row=0,column=1,padx=6,pady=6)

        ctk.CTkLabel(cfg, text="Key B:").grid(row=0,column=2,padx=6,pady=6, sticky="w")
        self.optKeyB = ctk.CTkOptionMenu(cfg, values=[], width=180)
        self.optKeyB.grid(row=0,column=3,padx=6,pady=6)

        ctk.CTkLabel(cfg, text="Cột A:").grid(row=1,column=0,padx=6,pady=6, sticky="w")
        self.optColA = ctk.CTkOptionMenu(cfg, values=[], width=180)
        self.optColA.grid(row=1,column=1,padx=6,pady=6)

        ctk.CTkLabel(cfg, text="Cột B:").grid(row=1,column=2,padx=6,pady=6, sticky="w")
        self.optColB = ctk.CTkOptionMenu(cfg, values=[], width=180)
        self.optColB.grid(row=1,column=3,padx=6,pady=6)

        ctk.CTkButton(cfg, text="+ Thêm cặp", command=self._add_pair).grid(row=1,column=4,padx=6)

        lower = ctk.CTkFrame(self)
        lower.pack(fill="both", expand=True, padx=pad, pady=6)

        left = ctk.CTkFrame(lower, width=420)
        left.pack(side="left", fill="y", padx=6, pady=6)

        ctk.CTkLabel(left, text="Các cặp (A ⇄ B)").pack(anchor="w")
        self.txtPairs = ctk.CTkTextbox(left, height=100)
        self.txtPairs.pack(fill="x", pady=6)

        ctk.CTkLabel(left, text="Chọn cột xuất từ File A").pack(anchor="w")
        self.frameColsA = ctk.CTkScrollableFrame(left, height=200)
        self.frameColsA.pack(fill="both", pady=6)

        ctk.CTkLabel(left, text="Chọn cột xuất từ File B").pack(anchor="w")
        self.frameColsB = ctk.CTkScrollableFrame(left, height=200)
        self.frameColsB.pack(fill="both", pady=6)

        right = ctk.CTkFrame(lower)
        right.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        # buttons row
        act = ctk.CTkFrame(right)
        act.pack(fill="x", pady=6)
        self.btnStart = ctk.CTkButton(act, text="▶ Bắt đầu", fg_color="green", command=self._start_worker)
        self.btnStart.pack(side="left", padx=6)
        self.btnPause = ctk.CTkButton(act, text="⏸ Tạm dừng", command=self._pause_resume, state="disabled")
        self.btnPause.pack(side="left", padx=6)
        self.btnStop = ctk.CTkButton(act, text="⏹ Dừng", fg_color="#ff5c5c", command=self._stop, state="disabled")
        self.btnStop.pack(side="left", padx=6)
        self.btnExport = ctk.CTkButton(act, text="💾 Xuất Excel", fg_color="#0b84ff", command=self._export, state="disabled")
        self.btnExport.pack(side="right", padx=6)

        self.progress = ttk.Progressbar(act, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=6)

        # preview tree
        self.tree = ttk.Treeview(right, columns=("kA","kB","status","detail"), show="headings")
        self.tree.heading("kA", text="Key A")
        self.tree.heading("kB", text="Key B")
        self.tree.heading("status", text="Trạng thái")
        self.tree.heading("detail", text="Chi tiết")
        self.tree.pack(fill="both", expand=True, side="left")

        vsb = ttk.Scrollbar(right, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscroll=vsb.set)

        # log
        self.txtLog = ctk.CTkTextbox(self, height=140)
        self.txtLog.pack(fill="x", padx=pad, pady=6)

    # ---------------- File selection and helpers ----------------
    def _choose_file_a(self):
        path = ctk.filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path: return
        self.fileA = path
        self.lblA.configure(text=f"File A: {path}")
        try:
            self.colsA = load_excel_columns(path)
            self.optKeyA.configure(values=self.colsA)
            self.optColA.configure(values=self.colsA)
            self._populate_cols(self.frameColsA, self.colsA)
            self._log(f"Load file A: {path} ({len(self.colsA)} cột)")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e), parent=self)

    def _choose_file_b(self):
        path = ctk.filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path: return
        self.fileB = path
        self.lblB.configure(text=f"File B: {path}")
        try:
            self.colsB = load_excel_columns(path)
            self.optKeyB.configure(values=self.colsB)
            self.optColB.configure(values=self.colsB)
            self._populate_cols(self.frameColsB, self.colsB)
            self._log(f"Load file B: {path} ({len(self.colsB)} cột)")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e), parent=self)

    def _populate_cols(self, frame, cols):
        for w in frame.winfo_children(): w.destroy()
        for c in cols:
            var = ctk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(frame, text=c, variable=var)
            chk.pack(anchor="w", pady=2)
            chk.var = var

    def _get_checked(self, frame):
        return [w.cget("text") for w in frame.winfo_children() if getattr(w,"var",None) and w.var.get()]

    def _add_pair(self):
        a = self.optColA.get(); b = self.optColB.get()
        if not a or not b:
            messagebox.showwarning("Thiếu", "Chọn cột A và B rồi nhấn Thêm", parent=self)
            return
        self.pairs.append((a,b))
        self.txtPairs.insert("end", f"{a} ⇄ {b}\n")
        self._log(f"Thêm: {a} ⇄ {b}")

    def _log(self, msg):
        self.txtLog.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n"); self.txtLog.see("end")

    # ---------------- theme change ----------------
    def _on_change_theme(self, *_):
        mode = self.appear_var.get()
        theme = self.theme_var.get()
        apply_style(mode, theme)
        # save config
        cfg = load_config()
        cfg['appearance_mode'] = mode
        cfg['color_theme'] = theme
        save_config(cfg)
        self._log(f"Đổi giao diện: {mode}, màu: {theme}")

    # ---------------- worker control ----------------
    def _start_worker(self):
        if not self.fileA or not self.fileB:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file A và file B", parent=self); return
        if not self.pairs:
            messagebox.showwarning("Thiếu cặp", "Thêm ít nhất 1 cặp A⇄B", parent=self); return
        keyA = self.optKeyA.get(); keyB = self.optKeyB.get()
        extraA = self._get_checked(self.frameColsA); extraB = self._get_checked(self.frameColsB)
        case_flag = False  # you can add checkbox for this
        remove_accents = False

        # disable buttons
        self.btnStart.configure(state="disabled"); self.btnPause.configure(state="normal"); self.btnStop.configure(state="normal"); self.btnExport.configure(state="disabled")
        self.stop_event.clear(); self.pause_event.clear()
        for i in self.tree.get_children(): self.tree.delete(i)
        self.progress['value'] = 0

        self.worker = threading.Thread(target=self._worker, args=(keyA,keyB,case_flag,remove_accents, extraA, extraB))
        self.worker.daemon = True
        self.worker.start()
        self._log("Bắt đầu so sánh...")

    def _pause_resume(self):
        if not self.pause_event.is_set():
            self.pause_event.set(); self.btnPause.configure(text="▶ Tiếp tục"); self._log("Tạm dừng")
        else:
            self.pause_event.clear(); self.btnPause.configure(text="⏸ Tạm dừng"); self._log("Tiếp tục")

    def _stop(self):
        self.stop_event.set(); self._log("Yêu cầu dừng...")

    def _worker(self, keyA, keyB, case_flag, remove_accents, extraA, extraB):
        try:
            df_result, preview = compare_tables(self.fileA, self.fileB, keyA, keyB, self.pairs, extraA, extraB, case_sensitive=case_flag, remove_accents=remove_accents)
            total = len(df_result)
            self.queue.put(("setmax", total))
            # iterate rows with stop/pause check; only push preview_limit rows to UI to stay snappy
            push_limit = 500
            pushed = 0
            for idx, row in df_result.iterrows():
                if self.stop_event.is_set():
                    self.queue.put(("stopped", None)); return
                while self.pause_event.is_set() and not self.stop_event.is_set():
                    time.sleep(0.2)
                kA = row.iloc[0] if len(row)>0 else ""
                kB = row.iloc[1] if len(row)>1 else ""
                status = row.get('Trạng thái',''); detail = row.get('Chi tiết','')
                if pushed < push_limit:
                    self.queue.put(("row", (kA,kB,status,detail))); pushed += 1
                self.queue.put(("progress", 1))
            self.queue.put(("done", df_result))
        except Exception as e:
            self.queue.put(("error", str(e)))

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                typ = item[0]
                if typ == "setmax":
                    self.progress['maximum'] = item[1]; self.progress['value'] = 0
                elif typ == "row":
                    kA,kB,status,detail = item[1]
                    iid = self.tree.insert("", "end", values=(kA,kB,status,detail))
                    if status == "Khớp":
                        self.tree.item(iid, tags=("match",)); self.tree.tag_configure("match", background="#e6f7e6")
                    elif status == "Chỉ bên A":
                        self.tree.item(iid, tags=("aonly",)); self.tree.tag_configure("aonly", background="#fff8dc")
                    else:
                        self.tree.item(iid, tags=("diff",)); self.tree.tag_configure("diff", background="#ffdede")
                elif typ == "progress":
                    self.progress['value'] += item[1]
                elif typ == "done":
                    df = item[1]; self.result_df = df
                    self._log(f"Hoàn tất: {len(df)} dòng.")
                    self.btnExport.configure(state="normal"); self.btnStart.configure(state="normal"); self.btnPause.configure(state="disabled"); self.btnStop.configure(state="disabled")
                elif typ == "stopped":
                    self._log("Đã dừng bởi người dùng."); self.btnStart.configure(state="normal"); self.btnPause.configure(state="disabled"); self.btnStop.configure(state="disabled")
                elif typ == "error":
                    self._log("Lỗi: "+item[1]); messagebox.showerror("Lỗi", item[1], parent=self); self.btnStart.configure(state="normal")
        except Exception:
            pass
        finally:
            self.after(100, self._process_queue)

    def _export(self):
        if self.result_df is None:
            messagebox.showwarning("Chưa có dữ liệu", "Chạy so sánh trước khi xuất", parent=self); return
        saved = save_result_dialog(self.result_df, parent=self)
        if saved:
            self._log(f"Đã lưu kết quả: {saved}")
            open_containing_folder(saved)
            # save config last
            cfg = load_config()
            cfg['last_saved'] = saved
            save_config(cfg)
