import os
from datetime import datetime

def ensure_folder(path):
    os.makedirs(path, exist_ok=True)
    return path

def timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")
