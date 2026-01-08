import os
import shutil

def safe_copy(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)

def is_xlsm_or_zip(path):
    return path.lower().endswith('.xlsm') or path.lower().endswith('.zip')

def is_exe(path):
    return path.lower().endswith('.exe')
