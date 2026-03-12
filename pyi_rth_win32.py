# PyInstaller runtime hook for pywin32
# Ensures pywin32 DLLs (pythoncom3xx.dll, pywintypes3xx.dll) are found at runtime.
import os
import sys

if getattr(sys, 'frozen', False):
    bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    if bundle_dir not in os.environ.get('PATH', ''):
        os.environ['PATH'] = bundle_dir + os.pathsep + os.environ.get('PATH', '')
