import os
import queue
import threading
import time
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.compare_engine import compare_tables
from core.excel_loader import load_excel_columns, load_excel_sheets, detect_header_row
from core.result_writer import save_result_dialog
from utils.helper import open_containing_folder


class CompareTab(ctk.CTkFrame):
    """Independent UI for report-style Excel comparison."""
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.file_a = self.file_b = None
        self.cols_a, self.cols_b, self.pairs = [], [], []
        self.result_df = None
        self.queue = queue.Queue()
        self._build()
        self.after(100, self._process_queue)

    def _build(self):
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=6, pady=6)
        ctk.CTkButton(top, text="Chọn file A", command=lambda: self._choose("a")).pack(side="left", padx=6, pady=6)
        self.lbl_a = ctk.CTkLabel(top, text="Chưa chọn file A", width=330, anchor="w")
        self.lbl_a.pack(side="left", padx=6)
        ctk.CTkButton(top, text="Chọn file B", command=lambda: self._choose("b")).pack(side="left", padx=6)
        self.lbl_b = ctk.CTkLabel(top, text="Chưa chọn file B", width=330, anchor="w")
        self.lbl_b.pack(side="left", padx=6)

        sheets = ctk.CTkFrame(self); sheets.pack(fill="x", padx=6, pady=3)
        ctk.CTkLabel(sheets, text="Sheet A").grid(row=0, column=0, padx=6, pady=6)
        self.sheet_a = ctk.CTkOptionMenu(sheets, values=[""], width=155, dynamic_resizing=False,
                                         command=lambda _: self._reload_side("a"))
        self.sheet_a.grid(row=0, column=1, padx=6)
        ctk.CTkLabel(sheets, text="Dòng tiêu đề A").grid(row=0, column=2, padx=6)
        self.header_a = ctk.CTkEntry(sheets, width=55); self.header_a.insert(0, "1")
        self.header_a.grid(row=0, column=3, padx=6)
        ctk.CTkLabel(sheets, text="Sheet B").grid(row=0, column=4, padx=6)
        self.sheet_b = ctk.CTkOptionMenu(sheets, values=[""], width=155, dynamic_resizing=False,
                                         command=lambda _: self._reload_side("b"))
        self.sheet_b.grid(row=0, column=5, padx=6)
        ctk.CTkLabel(sheets, text="Dòng tiêu đề B").grid(row=0, column=6, padx=6)
        self.header_b = ctk.CTkEntry(sheets, width=55); self.header_b.insert(0, "1")
        self.header_b.grid(row=0, column=7, padx=6)
        ctk.CTkButton(sheets, text="Tự tìm tiêu đề", command=self._auto_headers).grid(row=0, column=8, padx=6)
        ctk.CTkButton(sheets, text="Nạp lại cột", command=self._reload_both).grid(row=0, column=9, padx=6)

        mapping = ctk.CTkFrame(self)
        mapping.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(mapping, text="Khóa A").grid(row=0, column=0, padx=6, pady=6)
        self.key_a = ttk.Combobox(mapping, values=[""], width=27, height=18, state="readonly")
        self.key_a.grid(row=0, column=1, padx=6)
        ctk.CTkLabel(mapping, text="Khóa B").grid(row=0, column=2, padx=6)
        self.key_b = ttk.Combobox(mapping, values=[""], width=27, height=18, state="readonly")
        self.key_b.grid(row=0, column=3, padx=6)
        ctk.CTkLabel(mapping, text="Cột A").grid(row=1, column=0, padx=6, pady=6)
        self.col_a = ttk.Combobox(mapping, values=[""], width=27, height=18, state="readonly")
        self.col_a.grid(row=1, column=1, padx=6)
        ctk.CTkLabel(mapping, text="Cột B").grid(row=1, column=2, padx=6)
        self.col_b = ttk.Combobox(mapping, values=[""], width=27, height=18, state="readonly")
        self.col_b.grid(row=1, column=3, padx=6)
        ctk.CTkButton(mapping, text="+ Thêm cặp so sánh", command=self._add_pair).grid(row=1, column=4, padx=6)
        ctk.CTkButton(mapping, text="Xóa cặp cuối", command=self._remove_pair, fg_color="#a85d00").grid(row=1, column=5, padx=6)

        self.pair_box = ctk.CTkTextbox(self, height=90)
        self.pair_box.pack(fill="x", padx=6, pady=6)

        exports = ctk.CTkFrame(self); exports.pack(fill="x", padx=6, pady=3)
        export_a = ctk.CTkFrame(exports); export_a.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        export_b = ctk.CTkFrame(exports); export_b.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        self._build_export_selector(export_a, "Chọn cột xuất từ File A", "a")
        self._build_export_selector(export_b, "Chọn cột xuất từ File B", "b")
        actions = ctk.CTkFrame(self)
        actions.pack(fill="x", padx=6, pady=6)
        self.start_btn = ctk.CTkButton(actions, text="Bắt đầu so sánh", command=self._start, fg_color="green")
        self.start_btn.pack(side="left", padx=6, pady=6)
        self.export_btn = ctk.CTkButton(actions, text="Xuất báo cáo", command=self._export, state="disabled")
        self.export_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(actions, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)

        self.tree = ttk.Treeview(self, columns=("a", "b", "status", "detail"), show="headings")
        for col, text, width in (("a", "Khóa A", 150), ("b", "Khóa B", 150), ("status", "Trạng thái", 100), ("detail", "Chi tiết", 500)):
            self.tree.heading(col, text=text); self.tree.column(col, width=width)
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)
        self.log = ctk.CTkTextbox(self, height=90)
        self.log.pack(fill="x", padx=6, pady=6)

    def _build_export_selector(self, parent, title, side):
        head = ctk.CTkFrame(parent, fg_color="transparent"); head.pack(fill="x")
        ctk.CTkLabel(head, text=title).pack(side="left", padx=4)
        ctk.CTkButton(head, text="Chọn tất cả", width=85,
                      command=lambda: self._set_all_exports(side, True)).pack(side="right", padx=3)
        ctk.CTkButton(head, text="Bỏ chọn", width=70, fg_color="#6b7280",
                      command=lambda: self._set_all_exports(side, False)).pack(side="right", padx=3)
        frame = ctk.CTkScrollableFrame(parent, height=125)
        frame.pack(fill="x", padx=3, pady=4)
        if side == "a": self.export_frame_a = frame
        else: self.export_frame_b = frame

    def _populate_exports(self, side, columns):
        frame = self.export_frame_a if side == "a" else self.export_frame_b
        for widget in frame.winfo_children(): widget.destroy()
        for column in columns:
            var = ctk.BooleanVar(value=False)
            checkbox = ctk.CTkCheckBox(frame, text=column, variable=var)
            checkbox.pack(anchor="w", padx=3, pady=2)
            checkbox.column_name = column; checkbox.select_var = var

    def _set_all_exports(self, side, selected):
        frame = self.export_frame_a if side == "a" else self.export_frame_b
        for widget in frame.winfo_children():
            var = getattr(widget, "select_var", None)
            if var is not None: var.set(selected)

    def _selected_exports(self, side):
        frame = self.export_frame_a if side == "a" else self.export_frame_b
        return [widget.column_name for widget in frame.winfo_children()
                if getattr(widget, "select_var", None) is not None and widget.select_var.get()]

    def _choose(self, side):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Excel (.xlsx)", "*.xlsx")])
        if not path:
            return
        try:
            sheet_names = load_excel_sheets(path)
        except Exception as exc:
            messagebox.showerror("Lỗi đọc file", str(exc), parent=self); return
        if side == "a":
            self.file_a = path
            self.lbl_a.configure(text=os.path.basename(path))
            self.sheet_a.configure(values=sheet_names); self.sheet_a.set(sheet_names[0])
        else:
            self.file_b = path
            self.lbl_b.configure(text=os.path.basename(path))
            self.sheet_b.configure(values=sheet_names); self.sheet_b.set(sheet_names[0])
        self._reload_side(side)

    def _header_row(self, side):
        entry = self.header_a if side == "a" else self.header_b
        value = int(entry.get())
        if value < 1: raise ValueError("Dòng tiêu đề phải từ 1 trở lên.")
        return value

    def _reload_side(self, side):
        path = self.file_a if side == "a" else self.file_b
        if not path: return
        sheet = self.sheet_a.get() if side == "a" else self.sheet_b.get()
        try: cols = load_excel_columns(path, sheet, self._header_row(side))
        except Exception as exc:
            messagebox.showerror("Lỗi đọc cột", str(exc), parent=self); return
        if side == "a":
            self.cols_a = cols; menus = (self.key_a, self.col_a)
        else:
            self.cols_b = cols; menus = (self.key_b, self.col_b)
        for menu in menus:
            menu.configure(values=cols); menu.set(cols[0] if cols else "")
        self._populate_exports(side, cols)

    def _reload_both(self):
        self._reload_side("a"); self._reload_side("b")

    def _auto_headers(self):
        try:
            for side in ("a", "b"):
                path = self.file_a if side == "a" else self.file_b
                if not path: continue
                sheet = self.sheet_a.get() if side == "a" else self.sheet_b.get()
                row = detect_header_row(path, sheet)
                entry = self.header_a if side == "a" else self.header_b
                entry.delete(0, "end"); entry.insert(0, str(row))
            self._reload_both()
            self.log.insert("end", "Đã tự phát hiện dòng tiêu đề A/B.\n")
        except Exception as exc:
            messagebox.showerror("Lỗi tìm tiêu đề", str(exc), parent=self)

    def _add_pair(self):
        pair = (self.col_a.get(), self.col_b.get())
        if not all(pair): return
        self.pairs.append(pair); self._refresh_pairs()

    def _remove_pair(self):
        if self.pairs: self.pairs.pop(); self._refresh_pairs()

    def _refresh_pairs(self):
        self.pair_box.delete("1.0", "end")
        self.pair_box.insert("end", "\n".join(f"{a} ⇄ {b}" for a, b in self.pairs))

    def _start(self):
        if not self.file_a or not self.file_b or not self.pairs:
            messagebox.showwarning("Thiếu cấu hình", "Chọn hai file và thêm ít nhất một cặp cột.", parent=self); return
        missing_a = [c for pair in self.pairs for c in [pair[0]] if c not in self.cols_a]
        missing_b = [c for pair in self.pairs for c in [pair[1]] if c not in self.cols_b]
        if self.key_a.get() not in self.cols_a or self.key_b.get() not in self.cols_b or missing_a or missing_b:
            messagebox.showerror("Cấu hình không hợp lệ", "Sheet hoặc dòng tiêu đề đã thay đổi. Hãy nạp lại cột và chọn lại cấu hình.", parent=self); return
        config = (self.file_a, self.file_b, self.key_a.get(), self.key_b.get(), list(self.pairs),
                  self.sheet_a.get(), self.sheet_b.get(), self._header_row("a"), self._header_row("b"),
                  self._selected_exports("a"), self._selected_exports("b"))
        self.start_btn.configure(state="disabled"); self.export_btn.configure(state="disabled")
        self.progress.start(10)
        threading.Thread(target=self._worker, args=(config,), daemon=True).start()

    def _worker(self, config):
        try:
            file_a, file_b, key_a, key_b, pairs, sheet_a, sheet_b, header_a, header_b, extra_a, extra_b = config
            result, preview = compare_tables(file_a, file_b, key_a, key_b, pairs,
                                             extra_a=extra_a, extra_b=extra_b,
                                             sheet_a=sheet_a, sheet_b=sheet_b,
                                             header_row_a=header_a, header_row_b=header_b)
            self.queue.put(("done", result, preview))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                self.progress.stop(); self.start_btn.configure(state="normal")
                if item[0] == "error":
                    messagebox.showerror("Lỗi", item[1], parent=self)
                else:
                    self.result_df = item[1]
                    for child in self.tree.get_children(): self.tree.delete(child)
                    for _, row in item[2].iterrows():
                        vals = list(row.values)
                        self.tree.insert("", "end", values=((vals + ["", "", "", ""])[:2] + [row.get("Trạng thái", ""), row.get("Chi tiết", "")]))
                    self.export_btn.configure(state="normal")
                    self.log.insert("end", f"[{time.strftime('%H:%M:%S')}] Hoàn tất {len(self.result_df)} dòng.\n")
        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    def _export(self):
        saved = save_result_dialog(self.result_df, parent=self)
        if saved: open_containing_folder(saved)
