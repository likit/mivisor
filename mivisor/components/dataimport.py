import wx
import pandas as pd

class PreviewDialog(wx.Dialog):
    def __init__(self, parent):
        super(PreviewDialog, self).__init__(parent)
        self.dataSource = {} # for holding info about data source
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.filePathInput = wx.TextCtrl(self)
        browseButton = wx.Button(self, label='Browse')
        browseButton.Bind(wx.EVT_BUTTON, self.OnBrowse)
        hbox.Add(self.filePathInput, 0, wx.ALL | wx.EXPAND, 5)
        hbox.Add(browseButton)

        vbox = wx.BoxSizer(wx.VERTICAL)
        okBtn = wx.Button(self, wx.ID_OK)
        cancelBtn = wx.Button(self, wx.ID_CANCEL)
        vbox.Add(hbox, 0, wx.ALL|wx.EXPAND, 5)
        vbox.Add(okBtn, 0, wx.CENTER, 5)
        vbox.Add(cancelBtn, 0, wx.CENTER, 5)
        self.SetSizer(vbox)

    def OnBrowse(self, e):
        with wx.FileDialog(self, "Open data file",
                           wildcard="Excel files (*.xls;*xlsx)|*.xls;*.xlsx",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            filepath = fileDialog.GetPath()
            if filepath:
                self.filePathInput.SetValue(filepath)
