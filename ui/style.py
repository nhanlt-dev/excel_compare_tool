import customtkinter as ctk

def init_style(mode="system", theme="blue"):
    """
    mode: 'dark', 'light', 'system'
    theme: 'blue','green'...
    """
    ctk.set_appearance_mode(mode)
    ctk.set_default_color_theme(theme)
