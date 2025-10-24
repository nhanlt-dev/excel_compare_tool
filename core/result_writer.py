import os
import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox

def save_result_dialog(df, parent=None):
    default = f"compare_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = filedialog.asksaveasfilename(parent=parent,
                                        defaultextension=".xlsx",
                                        filetypes=[("Excel files", "*.xlsx")],
                                        initialfile=default,
                                        title="Chọn nơi lưu kết quả")
    if not path:
        return None
    # Try to use xlsxwriter for formatting
    try:
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="KQ")
            workbook = writer.book
            worksheet = writer.sheets["KQ"]
            # format rows by status if exists
            if "Trạng thái" in df.columns:
                fmt_match = workbook.add_format({'bg_color': '#dff0d8'})   # greenish
                fmt_diff  = workbook.add_format({'bg_color': '#ffd6d6'})   # red
                fmt_aonly = workbook.add_format({'bg_color': '#fff3cd'})   # yellow
                st_col = df.columns.get_loc("Trạng thái")
                for i, val in enumerate(df["Trạng thái"], start=1):
                    if val == "Khớp":
                        worksheet.set_row(i, None, fmt_match)
                    elif val == "Chỉ bên A":
                        worksheet.set_row(i, None, fmt_aonly)
                    else:
                        worksheet.set_row(i, None, fmt_diff)
        return os.path.abspath(path)
    except Exception as e:
        # fallback simple save
        try:
            df.to_excel(path, index=False)
            return os.path.abspath(path)
        except Exception as e2:
            messagebox.showerror("Lỗi lưu file", f"{e}\nFallback lỗi: {e2}", parent=parent)
            return None
