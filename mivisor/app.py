import wx
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

from components.main import MainWindow

def main():
    app = wx.App()
    mw = MainWindow(None)
    mw.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()