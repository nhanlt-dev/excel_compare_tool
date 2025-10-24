import pandas as pd
import os
from datetime import datetime
from tkinter import filedialog, messagebox

def save_result_dialog(df, parent=None):
    default = f"compare_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = filedialog.asksaveasfilename(parent=parent, defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")], initialfile=default, title="Chọn nơi lưu kết quả")
    if not path:
        return None
    try:
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Details', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Details']
            # Color rows by status
            if 'Trạng thái' in df.columns:
                fmt_match = workbook.add_format({'bg_color':'#e6f7e6'})
                fmt_diff  = workbook.add_format({'bg_color':'#fff0d9'})
                fmt_aonly = workbook.add_format({'bg_color':'#fff8dc'})
                for i, v in enumerate(df['Trạng thái'], start=1):
                    if v == 'Khớp':
                        worksheet.set_row(i, None, fmt_match)
                    elif v == 'Chỉ bên A':
                        worksheet.set_row(i, None, fmt_aonly)
                    else:
                        worksheet.set_row(i, None, fmt_diff)
            # summary sheet
            try:
                summary = df['Trạng thái'].value_counts(dropna=False).to_frame().reset_index()
                summary.columns = ['Trạng thái','Số lượng']
                summary.to_excel(writer, sheet_name='Summary', index=False)
            except Exception:
                pass
        return os.path.abspath(path)
    except Exception as e:
        try:
            df.to_excel(path, index=False)
            return os.path.abspath(path)
        except Exception as e2:
            messagebox.showerror("Lỗi lưu file", f"{e}\n{e2}", parent=parent)
            return None
