import os
import json
import queue
import threading
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.excel_loader import load_excel_columns, load_excel_sheets, detect_header_row
from core.lookup_update_engine import lookup_and_update, preview_lookup, undo_from_audit
from utils.helper import open_containing_folder


class LookupUpdateTab(ctk.CTkFrame):
    """Independent UI for lookup-and-write operations."""
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.source_path = self.target_path = None
        self.source_cols, self.target_cols = [], []
        self.key_pairs, self.value_pairs = [], []
        self.jobs = []
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self._build()
        self.after(100, self._process_queue)

    def _build(self):
        files = ctk.CTkFrame(self)
        files.pack(fill="x", padx=6, pady=6)
        ctk.CTkButton(files, text="Chọn file nguồn", command=lambda: self._choose_file("source")).grid(row=0, column=0, padx=6, pady=6)
        self.source_label = ctk.CTkLabel(files, text="Chưa chọn", width=300, anchor="w")
        self.source_label.grid(row=0, column=1, padx=6)
        ctk.CTkLabel(files, text="Sheet nguồn").grid(row=0, column=2, padx=6)
        self.source_sheet = ctk.CTkOptionMenu(
            files, values=[""], width=140, dynamic_resizing=False,
            command=lambda _: self._reload_columns("source")
        )
        self.source_sheet.grid(row=0, column=3, padx=6)
        ctk.CTkLabel(files, text="Dòng tiêu đề").grid(row=0, column=4, padx=6)
        self.source_header = ctk.CTkEntry(files, width=55); self.source_header.insert(0, "1")
        self.source_header.grid(row=0, column=5, padx=6)
        self.same_file_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(files, text="Nguồn và đích cùng file", variable=self.same_file_var,
                        command=self._toggle_same_file).grid(row=0, column=6, padx=8)

        ctk.CTkButton(files, text="Chọn file đích", command=lambda: self._choose_file("target")).grid(row=1, column=0, padx=6, pady=6)
        self.target_label = ctk.CTkLabel(files, text="Chưa chọn", width=300, anchor="w")
        self.target_label.grid(row=1, column=1, padx=6)
        ctk.CTkLabel(files, text="Sheet đích").grid(row=1, column=2, padx=6)
        self.target_sheet = ctk.CTkOptionMenu(
            files, values=[""], width=140, dynamic_resizing=False,
            command=lambda _: self._reload_columns("target")
        )
        self.target_sheet.grid(row=1, column=3, padx=6)
        ctk.CTkLabel(files, text="Dòng tiêu đề").grid(row=1, column=4, padx=6)
        self.target_header = ctk.CTkEntry(files, width=55); self.target_header.insert(0, "1")
        self.target_header.grid(row=1, column=5, padx=6)
        ctk.CTkButton(files, text="Tự tìm tiêu đề", command=self._auto_headers).grid(row=1, column=6, padx=6)
        ctk.CTkButton(files, text="Nạp lại cột", command=self._reload_both).grid(row=0, column=7, rowspan=2, padx=10)

        maps = ctk.CTkFrame(self)
        maps.pack(fill="both", expand=True, padx=6, pady=6)
        left = ctk.CTkFrame(maps); left.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        right = ctk.CTkFrame(maps); right.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self._build_mapping(left, "CẶP KHÓA ĐỐI CHIẾU", "key")
        self._build_mapping(right, "CỘT LẤY → CỘT GHI", "value")

        search = ctk.CTkFrame(self); search.pack(fill="x", padx=6, pady=2)
        ctk.CTkLabel(search, text="Tìm nhanh tên cột:").pack(side="left", padx=6)
        self.search_entry = ctk.CTkEntry(search, placeholder_text="Ví dụ: ketqua, ma_lk...", width=260)
        self.search_entry.pack(side="left", padx=5)
        ctk.CTkButton(search, text="Lọc", width=65, command=self._filter_columns).pack(side="left", padx=4)
        ctk.CTkButton(search, text="Hiện tất cả", width=90, command=self._reset_filter).pack(side="left", padx=4)
        ctk.CTkButton(search, text="Lưu cấu hình", width=95, fg_color="#0f766e", command=self._save_profile).pack(side="right", padx=4)
        ctk.CTkButton(search, text="Mở cấu hình", width=95, fg_color="#0f766e", command=self._load_profile).pack(side="right", padx=4)

        options = ctk.CTkFrame(self)
        options.pack(fill="x", padx=6, pady=6)
        self.duplicate = self._option(options, "Khóa nguồn trùng", ["Lấy dòng đầu", "Lấy dòng cuối", "Báo lỗi"], 0)
        self.missing = self._option(options, "Không tìm thấy", ["Giữ nguyên", "Xóa nội dung", "Ghi thông báo"], 2)
        self.overwrite = self._option(options, "Ô đích đã có dữ liệu", ["Ghi đè", "Chỉ điền ô trống"], 4)
        self.case_var = ctk.BooleanVar(value=False)
        self.accent_var = ctk.BooleanVar(value=False)
        self.highlight_var = ctk.BooleanVar(value=True)
        self.report_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options, text="Phân biệt hoa/thường", variable=self.case_var).grid(row=0, column=6, padx=8)
        ctk.CTkCheckBox(options, text="Bỏ dấu tiếng Việt khi dò", variable=self.accent_var).grid(row=0, column=7, padx=8)
        ctk.CTkCheckBox(options, text="Tô vàng ô cập nhật", variable=self.highlight_var).grid(row=1, column=0, columnspan=2, padx=8, pady=(0, 7), sticky="w")
        ctk.CTkCheckBox(options, text="Tạo các Sheet báo cáo phụ", variable=self.report_var).grid(row=1, column=2, columnspan=3, padx=8, pady=(0, 7), sticky="w")

        actions = ctk.CTkFrame(self)
        actions.pack(fill="x", padx=6, pady=6)
        self.start_btn = ctk.CTkButton(actions, text="Dò và ghi dữ liệu", fg_color="green", command=self._start)
        self.start_btn.pack(side="left", padx=6, pady=6)
        self.preview_btn = ctk.CTkButton(actions, text="Xem trước 20 dòng", command=self._preview)
        self.preview_btn.pack(side="left", padx=6)
        ctk.CTkButton(actions, text="Hoàn tác từ nhật ký", fg_color="#7c3aed",
                      command=self._undo).pack(side="left", padx=6)
        self.stop_btn = ctk.CTkButton(actions, text="Dừng", fg_color="#b33333", state="disabled", command=self.stop_event.set)
        self.stop_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(actions, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)
        batch = ctk.CTkFrame(self); batch.pack(fill="x", padx=6, pady=2)
        ctk.CTkButton(batch, text="+ Thêm tác vụ vào lô", command=self._add_job).pack(side="left", padx=6, pady=5)
        ctk.CTkButton(batch, text="Xóa tác vụ cuối", fg_color="#a85d00", command=self._remove_job).pack(side="left", padx=6)
        self.batch_btn = ctk.CTkButton(batch, text="Chạy lô (0 tác vụ)", fg_color="#2563eb", command=self._start_batch)
        self.batch_btn.pack(side="left", padx=6)
        self.job_label = ctk.CTkLabel(batch, text="Chưa có tác vụ trong lô", anchor="w")
        self.job_label.pack(side="left", fill="x", expand=True, padx=10)
        self.status = ctk.CTkLabel(self, text="Sẵn sàng", anchor="w")
        self.status.pack(fill="x", padx=12, pady=6)

    def _build_mapping(self, parent, title, kind):
        ctk.CTkLabel(parent, text=title, font=ctk.CTkFont(weight="bold")).pack(pady=5)
        row = ctk.CTkFrame(parent); row.pack(fill="x", padx=5)
        src = ttk.Combobox(
            row, values=[""], width=29, height=18, state="readonly"
        ); src.pack(side="left", padx=4, pady=5, ipady=4)
        ctk.CTkLabel(row, text="→").pack(side="left")
        dst = ttk.Combobox(
            row, values=[""], width=29, height=18, state="readonly"
        ); dst.pack(side="left", padx=4, ipady=4)
        ctk.CTkButton(row, text="Thêm", width=60, command=lambda: self._add_pair(kind)).pack(side="left", padx=3)
        ctk.CTkButton(row, text="Xóa cuối", width=70, fg_color="#a85d00", command=lambda: self._remove_pair(kind)).pack(side="left", padx=3)
        box = ctk.CTkTextbox(parent, height=160); box.pack(fill="both", expand=True, padx=5, pady=5)
        if kind == "key": self.key_source, self.key_target, self.key_box = src, dst, box
        else: self.value_source, self.value_target, self.value_box = src, dst, box

    def _option(self, parent, label, values, column):
        ctk.CTkLabel(parent, text=label).grid(row=0, column=column, padx=5, pady=8)
        menu = ctk.CTkOptionMenu(
            parent, values=values, width=145, dynamic_resizing=False
        )
        menu.grid(row=0, column=column + 1, padx=5)
        menu.set(values[0]); return menu

    def _choose_file(self, side):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Excel (.xlsx)", "*.xlsx")])
        if not path: return
        try:
            sheets = load_excel_sheets(path)
        except Exception as exc:
            messagebox.showerror("Lỗi đọc file", str(exc), parent=self); return
        menu = self.source_sheet if side == "source" else self.target_sheet
        menu.configure(values=sheets); menu.set(sheets[0])
        if side == "source":
            self.source_path = path; self.source_label.configure(text=os.path.basename(path))
            if self.same_file_var.get():
                self.target_path = path; self.target_label.configure(text=os.path.basename(path))
                self.target_sheet.configure(values=sheets); self.target_sheet.set(sheets[0])
        else:
            self.target_path = path; self.target_label.configure(text=os.path.basename(path))
        self._reload_columns(side)
        if side == "source" and self.same_file_var.get(): self._reload_columns("target")

    def _toggle_same_file(self):
        if self.same_file_var.get() and self.source_path:
            self.target_path = self.source_path
            self.target_label.configure(text=os.path.basename(self.source_path))
            sheets = load_excel_sheets(self.source_path)
            self.target_sheet.configure(values=sheets); self.target_sheet.set(sheets[0])
            self._reload_columns("target")

    def _auto_headers(self):
        try:
            for side in ("source", "target"):
                path = self.source_path if side == "source" else self.target_path
                if not path: continue
                sheet = self.source_sheet.get() if side == "source" else self.target_sheet.get()
                row = detect_header_row(path, sheet)
                entry = self.source_header if side == "source" else self.target_header
                entry.delete(0, "end"); entry.insert(0, str(row))
            self._reload_both(); self.status.configure(text="Đã tự phát hiện dòng tiêu đề.")
        except Exception as exc:
            messagebox.showerror("Lỗi tìm tiêu đề", str(exc), parent=self)

    def _filter_columns(self):
        term = self.search_entry.get().strip().casefold()
        if not term: return
        src = [c for c in self.source_cols if term in c.casefold()]
        dst = [c for c in self.target_cols if term in c.casefold()]
        for menu in (self.key_source, self.value_source): menu.configure(values=src or [""]); menu.set(src[0] if src else "")
        for menu in (self.key_target, self.value_target): menu.configure(values=dst or [""]); menu.set(dst[0] if dst else "")
        self.status.configure(text=f"Tìm thấy {len(src)} cột nguồn, {len(dst)} cột đích.")

    def _reset_filter(self):
        self.search_entry.delete(0, "end")
        for menu in (self.key_source, self.value_source): menu.configure(values=self.source_cols)
        for menu in (self.key_target, self.value_target): menu.configure(values=self.target_cols)

    def _save_profile(self):
        profile = {
            "version": 1,
            "source_sheet": self.source_sheet.get(), "target_sheet": self.target_sheet.get(),
            "source_header_row": self._header_row("source"), "target_header_row": self._header_row("target"),
            "key_pairs": self.key_pairs, "value_pairs": self.value_pairs,
            "duplicate": self.duplicate.get(), "missing": self.missing.get(),
            "overwrite": self.overwrite.get(), "case_sensitive": self.case_var.get(),
            "remove_accents": self.accent_var.get(), "highlight": self.highlight_var.get(),
            "create_reports": self.report_var.get(),
            "same_file": self.same_file_var.get()
        }
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".json",
                                             filetypes=[("Cấu hình JSON", "*.json")],
                                             initialfile="cau_hinh_excel_compare.json")
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(profile, handle, ensure_ascii=False, indent=2)
            self.status.configure(text="Đã lưu hồ sơ cấu hình.")
        except Exception as exc:
            messagebox.showerror("Lỗi lưu cấu hình", str(exc), parent=self)

    def _load_profile(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Cấu hình JSON", "*.json")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as handle: profile = json.load(handle)
            if profile.get("version") != 1: raise ValueError("Phiên bản cấu hình không được hỗ trợ.")
            self.same_file_var.set(bool(profile.get("same_file", False)))
            if self.same_file_var.get(): self._toggle_same_file()
            for entry, value in ((self.source_header, profile["source_header_row"]), (self.target_header, profile["target_header_row"])):
                entry.delete(0, "end"); entry.insert(0, str(value))
            if profile.get("source_sheet") in self.source_sheet.cget("values"): self.source_sheet.set(profile["source_sheet"])
            if profile.get("target_sheet") in self.target_sheet.cget("values"): self.target_sheet.set(profile["target_sheet"])
            self._reload_both()
            self.key_pairs = [tuple(x) for x in profile.get("key_pairs", [])]
            self.value_pairs = [tuple(x) for x in profile.get("value_pairs", [])]
            self.duplicate.set(profile.get("duplicate", "Lấy dòng đầu")); self.missing.set(profile.get("missing", "Giữ nguyên"))
            self.overwrite.set(profile.get("overwrite", "Ghi đè")); self.case_var.set(bool(profile.get("case_sensitive")))
            self.accent_var.set(bool(profile.get("remove_accents"))); self.highlight_var.set(bool(profile.get("highlight", True)))
            self.report_var.set(bool(profile.get("create_reports", False)))
            self._refresh("key"); self._refresh("value")
            self.status.configure(text="Đã mở hồ sơ cấu hình. Hãy kiểm tra và xem trước trước khi chạy.")
        except Exception as exc:
            messagebox.showerror("Lỗi mở cấu hình", str(exc), parent=self)

    def _header_row(self, side):
        widget = self.source_header if side == "source" else self.target_header
        value = int(widget.get())
        if value < 1: raise ValueError("Dòng tiêu đề phải từ 1 trở lên.")
        return value

    def _reload_columns(self, side):
        path = self.source_path if side == "source" else self.target_path
        if not path: return
        sheet = self.source_sheet.get() if side == "source" else self.target_sheet.get()
        try:
            cols = load_excel_columns(path, sheet, self._header_row(side))
        except Exception as exc:
            messagebox.showerror("Lỗi đọc cột", str(exc), parent=self); return
        if side == "source":
            self.source_cols = cols
            for menu in (self.key_source, self.value_source): menu.configure(values=cols); menu.set(cols[0] if cols else "")
        else:
            self.target_cols = cols
            for menu in (self.key_target, self.value_target): menu.configure(values=cols); menu.set(cols[0] if cols else "")

    def _reload_both(self):
        self._reload_columns("source"); self._reload_columns("target")

    def _add_pair(self, kind):
        pair = (self.key_source.get(), self.key_target.get()) if kind == "key" else (self.value_source.get(), self.value_target.get())
        if not all(pair): return
        pairs = self.key_pairs if kind == "key" else self.value_pairs
        if pair not in pairs: pairs.append(pair)
        self._refresh(kind)

    def _remove_pair(self, kind):
        pairs = self.key_pairs if kind == "key" else self.value_pairs
        if pairs: pairs.pop()
        self._refresh(kind)

    def _refresh(self, kind):
        pairs, box = (self.key_pairs, self.key_box) if kind == "key" else (self.value_pairs, self.value_box)
        box.delete("1.0", "end"); box.insert("end", "\n".join(f"{a} → {b}" for a, b in pairs))

    def _validate_config(self):
        if not self.source_path or not self.target_path or not self.key_pairs or not self.value_pairs:
            raise ValueError("Chọn đủ hai file, cặp khóa và cặp cột lấy/ghi.")
        warnings = []
        missing_source = sorted({a for a, _ in self.key_pairs + self.value_pairs if a not in self.source_cols})
        missing_target = sorted({b for _, b in self.key_pairs + self.value_pairs if b not in self.target_cols})
        if missing_source or missing_target:
            raise ValueError("Cấu hình không phù hợp file hiện tại. Thiếu cột nguồn: "
                             + ", ".join(missing_source or ["không"]) + "; thiếu cột đích: "
                             + ", ".join(missing_target or ["không"]))
        key_sources = {a.casefold() for a, _ in self.key_pairs}
        for source, target in self.value_pairs:
            if source.casefold() in key_sources:
                warnings.append(f"'{source}' vừa là khóa vừa là cột lấy.")
            if "tenbenh" in source.casefold() and "ketqua" in target.casefold():
                warnings.append(f"Có thể chọn nhầm: {source} → {target}.")
        return warnings

    def _preview(self):
        try:
            warnings = self._validate_config()
            rows = preview_lookup(
                self.source_path, self.target_path, self.source_sheet.get(), self.target_sheet.get(),
                self.key_pairs, self.value_pairs, self._header_row("source"), self._header_row("target"),
                self.case_var.get(), self.accent_var.get(), 20
            )
        except Exception as exc:
            messagebox.showerror("Không thể xem trước", str(exc), parent=self); return
        win = ctk.CTkToplevel(self); win.title("Xem trước - chưa ghi dữ liệu"); win.geometry("1000x520")
        if warnings:
            ctk.CTkLabel(win, text="CẢNH BÁO: " + " | ".join(warnings), text_color="#d97706", wraplength=950).pack(fill="x", padx=8, pady=6)
        tree = ttk.Treeview(win, columns=("row","key","old","new","status"), show="headings")
        for col, title, width in (("row","Dòng",60),("key","Khóa",280),("old","Giá trị cũ",210),("new","Giá trị mới",210),("status","Trạng thái",110)):
            tree.heading(col, text=title); tree.column(col, width=width)
        for row in rows: tree.insert("", "end", values=(row["row"], row["key"], row["old"], row["new"], row["status"]))
        tree.pack(fill="both", expand=True, padx=8, pady=8)

    def _start(self):
        try:
            warnings = self._validate_config()
        except ValueError as exc:
            messagebox.showwarning("Thiếu cấu hình", str(exc), parent=self); return
        summary = "Khóa:\n" + "\n".join(f"  {a} → {b}" for a, b in self.key_pairs)
        summary += "\n\nCột lấy → ghi:\n" + "\n".join(f"  {a} → {b}" for a, b in self.value_pairs)
        if warnings: summary += "\n\nCẢNH BÁO:\n" + "\n".join(warnings)
        if not messagebox.askyesno("Xác nhận cấu hình", summary + "\n\nTiếp tục chạy?", parent=self): return
        output = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="ket_qua_cap_nhat.xlsx")
        if not output: return
        # Read every Tkinter value on the main thread. Calling widget.get()
        # from the worker can freeze Tkinter on some Windows machines.
        config = {
            "source_path": self.source_path,
            "target_path": self.target_path,
            "source_sheet": self.source_sheet.get(),
            "target_sheet": self.target_sheet.get(),
            "key_pairs": list(self.key_pairs),
            "value_pairs": list(self.value_pairs),
            "source_header_row": self._header_row("source"),
            "target_header_row": self._header_row("target"),
            "duplicate_policy": {"Lấy dòng đầu": "first", "Lấy dòng cuối": "last", "Báo lỗi": "error"}[self.duplicate.get()],
            "missing_policy": {"Giữ nguyên": "keep", "Xóa nội dung": "clear", "Ghi thông báo": "text"}[self.missing.get()],
            "overwrite_policy": {"Ghi đè": "all", "Chỉ điền ô trống": "blank_only"}[self.overwrite.get()],
            "case_sensitive": self.case_var.get(),
            "remove_accents": self.accent_var.get(),
            "highlight_updates": self.highlight_var.get(),
            "create_audit": self.report_var.get(),
        }
        self.stop_event.clear(); self.start_btn.configure(state="disabled"); self.stop_btn.configure(state="normal")
        self.progress["value"] = 0; self.status.configure(text="Đang xử lý...")
        threading.Thread(target=self._worker, args=(output, config), daemon=True).start()

    def _worker(self, output, config):
        try:
            stats = lookup_and_update(
                config["source_path"], config["target_path"], output,
                config["source_sheet"], config["target_sheet"],
                config["key_pairs"], config["value_pairs"],
                config["source_header_row"], config["target_header_row"],
                config["duplicate_policy"], config["missing_policy"],
                config["overwrite_policy"],
                case_sensitive=config["case_sensitive"],
                remove_accents=config["remove_accents"],
                create_audit=config["create_audit"],
                highlight_updates=config["highlight_updates"],
                progress_callback=lambda phase, done, total: self.queue.put(
                    ("progress", phase, done, total)
                ),
                stop_event=self.stop_event
            )
            self.queue.put(("done", output, stats))
        except InterruptedError:
            self.queue.put(("stopped",))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _capture_job(self):
        self._validate_config()
        return {
            "source_path": self.source_path, "target_path": self.target_path,
            "source_sheet": self.source_sheet.get(), "target_sheet": self.target_sheet.get(),
            "key_pairs": list(self.key_pairs), "value_pairs": list(self.value_pairs),
            "source_header_row": self._header_row("source"), "target_header_row": self._header_row("target"),
            "duplicate_policy": {"Lấy dòng đầu":"first","Lấy dòng cuối":"last","Báo lỗi":"error"}[self.duplicate.get()],
            "missing_policy": {"Giữ nguyên":"keep","Xóa nội dung":"clear","Ghi thông báo":"text"}[self.missing.get()],
            "overwrite_policy": {"Ghi đè":"all","Chỉ điền ô trống":"blank_only"}[self.overwrite.get()],
            "case_sensitive": self.case_var.get(), "remove_accents": self.accent_var.get(),
            "highlight_updates": self.highlight_var.get(), "create_audit": self.report_var.get()
        }

    def _add_job(self):
        try: job = self._capture_job()
        except Exception as exc:
            messagebox.showwarning("Không thể thêm tác vụ", str(exc), parent=self); return
        self.jobs.append(job); self._refresh_jobs()

    def _remove_job(self):
        if self.jobs: self.jobs.pop()
        self._refresh_jobs()

    def _refresh_jobs(self):
        self.batch_btn.configure(text=f"Chạy lô ({len(self.jobs)} tác vụ)")
        if not self.jobs: self.job_label.configure(text="Chưa có tác vụ trong lô"); return
        last = self.jobs[-1]
        self.job_label.configure(text=f"Tác vụ cuối: {os.path.basename(last['source_path'])}/{last['source_sheet']} → {os.path.basename(last['target_path'])}/{last['target_sheet']}")

    def _start_batch(self):
        if not self.jobs:
            messagebox.showwarning("Chưa có tác vụ", "Hãy thêm ít nhất một tác vụ vào lô.", parent=self); return
        folder = filedialog.askdirectory(parent=self, title="Chọn thư mục lưu kết quả lô")
        if not folder: return
        self.stop_event.clear(); self.start_btn.configure(state="disabled"); self.batch_btn.configure(state="disabled"); self.stop_btn.configure(state="normal")
        threading.Thread(target=self._batch_worker, args=(folder, list(self.jobs)), daemon=True).start()

    def _batch_worker(self, folder, jobs):
        outputs=[]
        try:
            for number, job in enumerate(jobs, start=1):
                if self.stop_event.is_set(): raise InterruptedError()
                stem = os.path.splitext(os.path.basename(job["target_path"]))[0]
                output = os.path.join(folder, f"{stem}_cap_nhat_{number:03}.xlsx")
                self.queue.put(("batch_status", number, len(jobs), os.path.basename(output)))
                lookup_and_update(
                    job["source_path"], job["target_path"], output, job["source_sheet"], job["target_sheet"],
                    job["key_pairs"], job["value_pairs"], job["source_header_row"], job["target_header_row"],
                    job["duplicate_policy"], job["missing_policy"], job["overwrite_policy"],
                    case_sensitive=job["case_sensitive"], remove_accents=job["remove_accents"],
                    stop_event=self.stop_event, create_audit=job["create_audit"], highlight_updates=job["highlight_updates"]
                ); outputs.append(output)
            self.queue.put(("batch_done", folder, len(outputs)))
        except InterruptedError: self.queue.put(("stopped",))
        except Exception as exc: self.queue.put(("error", str(exc)))

    def _undo(self):
        updated = filedialog.askopenfilename(parent=self, title="Chọn file đã cập nhật",
                                              filetypes=[("Excel", "*.xlsx")])
        if not updated: return
        output = filedialog.asksaveasfilename(parent=self, title="Lưu file đã hoàn tác",
                                               defaultextension=".xlsx",
                                               filetypes=[("Excel", "*.xlsx")],
                                               initialfile="ket_qua_hoan_tac.xlsx")
        if not output: return
        try:
            restored = undo_from_audit(updated, output)
            messagebox.showinfo("Hoàn tác hoàn tất", f"Đã khôi phục {restored} ô vào file mới.", parent=self)
            open_containing_folder(output)
        except Exception as exc:
            messagebox.showerror("Không thể hoàn tác", str(exc), parent=self)

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if item[0] == "progress":
                    phase, done, total = item[1], item[2], item[3]
                    self.progress["maximum"] = max(1, total); self.progress["value"] = done
                    labels = {
                        "source": "Đang đọc và lập chỉ mục Sheet nguồn",
                        "target": "Đang dò và cập nhật Sheet đích",
                        "saving": "Đang lưu file kết quả",
                    }
                    percent = int(done * 100 / max(1, total))
                    self.status.configure(text=f"{labels.get(phase, 'Đang xử lý')}... {percent}%")
                elif item[0] == "stopped":
                    self.start_btn.configure(state="normal"); self.batch_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    self.status.configure(text="Đã dừng. File kết quả chưa hoàn chỉnh đã được xóa.")
                elif item[0] == "error":
                    self.start_btn.configure(state="normal"); self.batch_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    self.status.configure(text="Có lỗi: " + item[1]); messagebox.showerror("Lỗi", item[1], parent=self)
                elif item[0] == "batch_status":
                    self.status.configure(text=f"Đang chạy tác vụ {item[1]}/{item[2]}: {item[3]}")
                elif item[0] == "batch_done":
                    self.start_btn.configure(state="normal"); self.batch_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    self.status.configure(text=f"Hoàn tất {item[2]} tác vụ.")
                    messagebox.showinfo("Hoàn tất chạy lô", f"Đã tạo {item[2]} file kết quả.", parent=self); open_containing_folder(os.path.join(item[1], "dummy.xlsx"))
                else:
                    output, s = item[1], item[2]
                    self.start_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    text = (f"Hoàn tất: {s['matched']} khớp, {s['missing']} không tìm thấy, "
                            f"{s['updated_cells']} ô đã ghi, {s['duplicate_keys']} khóa nguồn trùng.")
                    self.status.configure(text=text)
                    messagebox.showinfo("Hoàn tất", text, parent=self); open_containing_folder(output)
        except queue.Empty:
            pass
        self.after(100, self._process_queue)
