import pandas as pd
import os
from datetime import datetime
from tkinter import filedialog, messagebox
from utils.logger import get_logger

logger = get_logger()

def save_result_dialog(df, parent=None):
    logger.info(f"Exporting result: {len(df)} rows")
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
                logger.warning("Failed to create summary sheet", exc_info=True)
                pass
        abs_path = os.path.abspath(path)
        logger.info(f"Result saved successfully: {abs_path}")
        return abs_path
    except Exception as e:
        logger.error(f"Failed to save Excel with xlsxwriter: {e}", exc_info=True)
        try:
            df.to_excel(path, index=False)
            abs_path = os.path.abspath(path)
            logger.info(f"Result saved with basic method: {abs_path}")
            return abs_path
        except Exception as e2:
            logger.error(f"Both export methods failed. Second error: {e2}", exc_info=True)
            messagebox.showerror("Lỗi lưu file", f"{e}\n{e2}", parent=parent)
            return None
