import sys
import customtkinter as ctk

from ui.compare_tab import CompareTab
from ui.lookup_update_tab import LookupUpdateTab
from ui.style import apply_style
from utils.config import load_config
from utils.logger import log_exception


class ExcelCompareApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        cfg = load_config()
        apply_style(cfg.get("appearance_mode", "system"), cfg.get("color_theme", "blue"))
        self.title("Excel Compare PRO 7.1.3")
        self.geometry("1280x820")
        self.minsize(1080, 700)

        tabs = ctk.CTkTabview(self)
        tabs.pack(fill="both", expand=True, padx=10, pady=10)
        compare = tabs.add("SO SÁNH DỮ LIỆU")
        lookup = tabs.add("DÒ VÀ GHI DỮ LIỆU")
        CompareTab(compare).pack(fill="both", expand=True)
        LookupUpdateTab(lookup).pack(fill="both", expand=True)


sys.excepthook = log_exception
