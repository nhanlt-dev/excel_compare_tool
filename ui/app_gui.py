import customtkinter as ctk
from tkinter import ttk, messagebox
import threading, queue, time
from core.excel_loader import load_headers, read_table, load_headers as load_excel_columns
from core.compare_engine import compare_tables
from core.result_writer import save_result_dialog
from utils.helper import ensure_folder, timestamp
from utils.config import save_config, load_config
from ui.style import init_style

import pandas as pd

init_style()

PREVIEW_LIMIT = 300  # số dòng preview hiển thị

class ExcelCompareApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("So sánh Excel — PRO 3.0")
        self.geometry("1200x760")

        # state
        self.fileA = None
        self.fileB = None
        self.colsA = []
        self.colsB = []
        self.pairs = []
        self.case_sensitive = False

        # thread control
        self.worker_thread = None
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()  # when set => paused

        self.result_df = None

        # load last config
        self.cfg = load_config()

        self._build_ui()
        self.after(100, self._process_queue)

    def _build_ui(self):
        pad = 10

        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=pad, pady=8)

        btnA = ctk.CTkButton(top, text="Chọn file A", width=140, command=self._choose_file_a)
        btnA.pack(side="left", padx=6)
        self.lblA = ctk.CTkLabel(top, text="Chưa chọn file A")
        self.lblA.pack(side="left", padx=6)

        btnB = ctk.CTkButton(top, text="Chọn file B", width=140, command=self._choose_file_b)
        btnB.pack(side="left", padx=6)
        self.lblB = ctk.CTkLabel(top, text="Chưa chọn file B")
        self.lblB.pack(side="left", padx=6)

        # Case checkbox
        self.case_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(top, text="Bỏ phân biệt hoa thường", variable=self.case_var).pack(side="right", padx=6)

        # mapping frame
        cfg = ctk.CTkFrame(self)
        cfg.pack(fill="x", padx=pad, pady=6)

        ctk.CTkLabel(cfg, text="Key A:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.cmbKeyA = ctk.CTkOptionMenu(cfg, values=[])
        self.cmbKeyA.grid(row=0, column=1, padx=6, pady=6)
        ctk.CTkLabel(cfg, text="Key B:").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        self.cmbKeyB = ctk.CTkOptionMenu(cfg, values=[])
        self.cmbKeyB.grid(row=0, column=3, padx=6, pady=6)

        ctk.CTkLabel(cfg, text="Cột A:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.cmbColA = ctk.CTkOptionMenu(cfg, values=[])
        self.cmbColA.grid(row=1, column=1, padx=6, pady=6)
        ctk.CTkLabel(cfg, text="Cột B:").grid(row=1, column=2, sticky="w", padx=6, pady=6)
        self.cmbColB = ctk.CTkOptionMenu(cfg, values=[])
        self.cmbColB.grid(row=1, column=3, padx=6, pady=6)
        ctk.CTkButton(cfg, text="+ Thêm cặp", command=self._add_pair).grid(row=1, column=4, padx=6)

        # pairs + columns selection
        lower = ctk.CTkFrame(self)
        lower.pack(fill="both", expand=False, padx=pad, pady=6)

        left = ctk.CTkFrame(lower)
        left.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(left, text="Các cặp so sánh (A ⇄ B)").pack(anchor="w")
        self.txtPairs = ctk.CTkTextbox(left, height=80)
        self.txtPairs.pack(fill="x", pady=6)

        ctk.CTkLabel(left, text="Chọn cột xuất từ File A").pack(anchor="w")
        self.frameColsA = ctk.CTkScrollableFrame(left, height=220)
        self.frameColsA.pack(fill="both", expand=True, pady=6)

        right = ctk.CTkFrame(lower)
        right.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(right, text="Chọn cột xuất từ File B").pack(anchor="w")
        self.frameColsB = ctk.CTkScrollableFrame(right, height=300)
        self.frameColsB.pack(fill="both", expand=True, pady=6)

        # action row
        act = ctk.CTkFrame(self)
        act.pack(fill="x", padx=pad, pady=6)

        self.btnStart = ctk.CTkButton(act, text="▶ Bắt đầu so sánh", fg_color="green", command=self._start_worker)
        self.btnStart.pack(side="left", padx=6)
        self.btnPause = ctk.CTkButton(act, text="⏸️ Tạm dừng", command=self._pause_resume, state="disabled")
        self.btnPause.pack(side="left", padx=6)
        self.btnStop = ctk.CTkButton(act, text="⏹ Dừng", fg_color="#ff5c5c", command=self._stop, state="disabled")
        self.btnStop.pack(side="left", padx=6)
        self.btnExport = ctk.CTkButton(act, text="💾 Xuất Excel", fg_color="#0b84ff", command=self._export, state="disabled")
        self.btnExport.pack(side="right", padx=6)

        self.progress = ttk.Progressbar(act, orient="horizontal", mode="determinate", length=600)
        self.progress.pack(fill="x", padx=6, pady=6)

        # preview tree
        preview = ctk.CTkFrame(self)
        preview.pack(fill="both", expand=True, padx=pad, pady=6)

        self.tree = ttk.Treeview(preview, columns=("kA","kB","status","detail"), show="headings")
        self.tree.heading("kA", text="Key A")
        self.tree.heading("kB", text="Key B")
        self.tree.heading("status", text="Trạng thái")
        self.tree.heading("detail", text="Chi tiết")
        self.tree.pack(fill="both", expand=True, side="left")

        vsb = ttk.Scrollbar(preview, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscroll=vsb.set)

        # log
        self.txtLog = ctk.CTkTextbox(self, height=140)
        self.txtLog.pack(fill="x", padx=pad, pady=6)

    # ---------------- File selection and UI helpers ----------------
    def _choose_file_a(self):
        path = ctk.filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path:
            return
        self.fileA = path
        self.lblA.configure(text=f"File A: {path}")
        try:
            self.colsA = load_excel_columns(path)
            self.cmbKeyA.configure(values=self.colsA)
            self.cmbColA.configure(values=self.colsA)
            self._populate_cols(self.frameColsA, self.colsA)
            self._log(f"Đã load file A: {path} ({len(self.colsA)} cột)")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file A", str(e), parent=self)

    def _choose_file_b(self):
        path = ctk.filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path:
            return
        self.fileB = path
        self.lblB.configure(text=f"File B: {path}")
        try:
            self.colsB = load_excel_columns(path)
            self.cmbKeyB.configure(values=self.colsB)
            self.cmbColB.configure(values=self.colsB)
            self._populate_cols(self.frameColsB, self.colsB)
            self._log(f"Đã load file B: {path} ({len(self.colsB)} cột)")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file B", str(e), parent=self)

    def _populate_cols(self, frame, cols):
        for w in frame.winfo_children():
            w.destroy()
        for c in cols:
            var = ctk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(frame, text=c, variable=var)
            chk.pack(anchor="w", pady=2, padx=6)
            chk.var = var

    def _get_checked(self, frame):
        return [w.cget("text") for w in frame.winfo_children() if getattr(w,"var",None) and w.var.get()]

    def _add_pair(self):
        a = self.cmbColA.get()
        b = self.cmbColB.get()
        if not a or not b:
            messagebox.showwarning("Thiếu", "Chọn cả cột A và cột B để thêm cặp", parent=self)
            return
        self.pairs.append((a,b))
        self.txtPairs.insert("end", f"{a} ⇄ {b}\n")
        self._log(f"Thêm cặp: {a} ⇄ {b}")

    def _log(self, msg):
        t = time.strftime("%H:%M:%S")
        self.txtLog.insert("end", f"[{t}] {msg}\n")
        self.txtLog.see("end")

    # ---------------- Thread control ----------------
    def _start_worker(self):
        if not self.fileA or not self.fileB:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file A và file B", parent=self)
            return
        if not self.pairs:
            messagebox.showwarning("Thiếu cặp", "Vui lòng thêm ít nhất 1 cặp để so sánh", parent=self)
            return

        # collect settings
        keyA = self.cmbKeyA.get()
        keyB = self.cmbKeyB.get()
        extraA = self._get_checked(self.frameColsA)
        extraB = self._get_checked(self.frameColsB)
        case_flag = not self.case_var.get()  # checkbox means "bỏ phân biệt", so case_sensitive = not checked

        # disable buttons
        self.btnStart.configure(state="disabled")
        self.btnPause.configure(state="normal")
        self.btnStop.configure(state="normal")
        self.btnExport.configure(state="disabled")

        # reset events
        self.stop_event.clear()
        self.pause_event.clear()

        # clear tree
        for i in self.tree.get_children():
            self.tree.delete(i)

        # launch thread
        self.worker_thread = threading.Thread(target=self._worker, args=(keyA,keyB,case_flag,extraA,extraB))
        self.worker_thread.daemon = True
        self.worker_thread.start()
        self._log("Worker thread started.")

    def _pause_resume(self):
        if not self.pause_event.is_set():
            # pause
            self.pause_event.set()
            self.btnPause.configure(text="▶ Tiếp tục")
            self._log("Paused.")
        else:
            # resume
            self.pause_event.clear()
            self.btnPause.configure(text="⏸️ Tạm dừng")
            self._log("Resumed.")

    def _stop(self):
        self.stop_event.set()
        self._log("Yêu cầu dừng...")

    def _worker(self, keyA, keyB, case_flag, extraA, extraB):
        try:
            # read full and compare (fast merge)
            df_result, preview = compare_tables(self.fileA, self.fileB, keyA, keyB, self.pairs, extraA, extraB, case_sensitive=case_flag, preview_limit=PREVIEW_LIMIT)
            # push preview rows to queue for UI
            total = len(df_result)
            self.queue.put(("setmax", total))
            for idx, row in df_result.iterrows():
                # pause handling
                while self.pause_event.is_set() and not self.stop_event.is_set():
                    time.sleep(0.2)
                if self.stop_event.is_set():
                    self.queue.put(("stopped", None))
                    return
                kA = row.iloc[0] if len(row)>0 else ""
                kB = row.iloc[1] if len(row)>1 else ""
                status = row.get("Trạng thái","")
                detail = row.get("Chi tiết","")
                self.queue.put(("row", (kA, kB, status, detail)))
                self.queue.put(("progress", 1))
                # limit UI push to PREVIEW_LIMIT to avoid huge UI cost
                # we still process full df_result, but only show first PREVIEW_LIMIT rows
                # here we continue pushing for all rows but UI may ignore beyond preview limit if desired
            self.queue.put(("done", df_result))
        except Exception as e:
            self.queue.put(("error", str(e)))

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                typ = item[0]
                if typ == "setmax":
                    total = item[1]
                    self.progress["maximum"] = total
                    self.progress["value"] = 0
                elif typ == "row":
                    kA,kB,status,detail = item[1]
                    iid = self.tree.insert("", "end", values=(kA,kB,status,detail))
                    # color by status
                    if status == "Khớp":
                        self.tree.item(iid, tags=("match",))
                        self.tree.tag_configure("match", background="#e9f7ea")
                    elif status == "Chỉ bên A":
                        self.tree.item(iid, tags=("aonly",))
                        self.tree.tag_configure("aonly", background="#fff8dc")
                    else:
                        self.tree.item(iid, tags=("diff",))
                        self.tree.tag_configure("diff", background="#ffdede")
                elif typ == "progress":
                    self.progress["value"] += item[1]
                elif typ == "done":
                    df = item[1]
                    self.result_df = df
                    self._log(f"Hoàn tất: {len(df)} dòng.")
                    # enable export
                    self.btnExport.configure(state="normal")
                    self.btnStart.configure(state="normal")
                    self.btnPause.configure(state="disabled")
                    self.btnStop.configure(state="disabled")
                elif typ == "stopped":
                    self._log("Đã dừng bởi người dùng.")
                    self.btnStart.configure(state="normal")
                    self.btnPause.configure(state="disabled")
                    self.btnStop.configure(state="disabled")
                elif typ == "error":
                    self._log("Lỗi trong worker: " + item[1])
                    messagebox.showerror("Lỗi", item[1], parent=self)
                    self.btnStart.configure(state="normal")
                    self.btnPause.configure(state="disabled")
                    self.btnStop.configure(state="disabled")
        except queue.Empty:
            pass
        finally:
            self.after(100, self._process_queue)

    def _export(self):
        if self.result_df is None:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng chạy so sánh trước.", parent=self)
            return
        saved = save_result_dialog(self.result_df, parent=self)
        if saved:
            self._log(f"Đã lưu file kết quả: {saved}")

