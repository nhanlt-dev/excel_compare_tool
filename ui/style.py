import customtkinter as ctk

# available modes: "dark","light","system"
# available accent themes depend on customtkinter: "blue","green","dark-blue", etc.

def apply_style(mode="system", theme="blue"):
    try:
        ctk.set_appearance_mode(mode)
        ctk.set_default_color_theme(theme)
    except Exception:
        # fallback default
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
