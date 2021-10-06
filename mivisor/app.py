import os
import sys

import wx
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

from components.main import GenApp


def main():
    app = GenApp()
    app.MainLoop()


if __name__ == '__main__':
    main()
