import os
from datetime import datetime
import webbrowser
def ensure_folder(p):
    os.makedirs(p, exist_ok=True)
    return p

def timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def open_containing_folder(path):
    if not path: return
    folder = os.path.dirname(path)
    try:
        webbrowser.open(folder)
    except Exception:
        pass
