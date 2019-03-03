import wx

from components.main import MainWindow

def main():
    app = wx.App()
    mw = MainWindow(None)
    mw.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()