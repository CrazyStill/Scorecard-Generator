"""
app_entry.py - Desktop application entry point.

Passes the Flask WSGI app directly to pywebview, which manages the internal
server. No threads, no port management, no polling required.

Usage (development):  python app_entry.py
Usage (production):   ScorecardCreator.exe  (built via PyInstaller)
"""
import sys
import os

if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    if bundle_dir not in sys.path:
        sys.path.insert(0, bundle_dir)

import webview
from main import app


def main():
    webview.create_window(
        title='Scorecard Creator',
        url=app,          # pass Flask WSGI app directly - pywebview serves it
        width=1100,
        height=780,
        min_size=(900, 600),
        maximized=True,
    )
    webview.start(gui='edgechromium', debug=False)


if __name__ == '__main__':
    main()
