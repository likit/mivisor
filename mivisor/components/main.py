import wx
import pandas
import xlrd
from components.datatable import DataGrid

def browse(filetype='MLAB'):
    file_meta = {
        'MLAB': {
            'wildcard': "Excel files (*.xls;*xlsx)|*.xls;*.xlsx"
        },
        'CSV': {
            'wildcard': "CSV files (*.csv)|*.csv"
        }
    }
    with wx.FileDialog(None, "Open data file",
                       wildcard=file_meta[filetype]['wildcard'],
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) \
            as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return

        return fileDialog.GetPath()


def show_sheets(parent, worksheets):
    dlg = wx.SingleChoiceDialog(None,
            "Select a worksheet", "Worksheets", worksheets)
    if dlg.ShowModal() == wx.ID_OK:
        return dlg.GetStringSelection()
    dlg.Destroy()


class MainWindow(wx.Frame):
    def __init__(self, parent):
        super(MainWindow, self).__init__(parent)
        self.SetTitle('Mivisor Version 1.0')
        self.SetSize((1200, 800))
        self.Center()

        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        imp = wx.Menu()
        mlabItem = imp.Append(wx.ID_ANY, 'MLAB')
        csvItem = imp.Append(wx.ID_ANY, 'CSV')
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_ANY, 'I&mport', imp)
        fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit Application')
        menubar.Append(fileMenu, '&File')
        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.OnQuit, fileItem)
        self.Bind(wx.EVT_MENU, self.OnLoadMLAB, mlabItem)
        self.Bind(wx.EVT_MENU, self.OnLoadCSV, csvItem)

        self.panel = wx.Panel(self, wx.ID_ANY)

        self.data_grid = DataGrid(self.panel)

        info_box = wx.StaticBox(self.panel, -1, 'Field Information')
        self.info_box_sizer = wx.StaticBoxSizer(info_box, wx.VERTICAL)
        lbl = wx.StaticText(info_box, label="Field info here")
        self.info_box_sizer.Add(lbl)

        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.vbox.Add(self.data_grid, 1, wx.EXPAND, 5)
        self.vbox.Add(self.info_box_sizer, 1, wx.ALL, 10)
        self.panel.SetSizer(self.vbox)


    def OnQuit(self, e):
        self.Close()

    def OnLoadMLAB(self, e):
        filepath = browse('MLAB')
        if filepath:
            try:
                worksheets = xlrd.open_workbook(filepath).sheet_names()
            except FileNotFoundError:
                wx.MessageDialog(self, 'Cannot download the data file.\nPlease check the file path again.',
                            'File Not Found!', wx.OK|wx.CENTER).ShowModal()
            else:
                if len(worksheets) > 1:
                    sel_worksheet = show_sheets(self, worksheets)
                else:
                    sel_worksheet = worksheets[0]
                df = pandas.read_excel(filepath, sheet_name=sel_worksheet)
                self.data_grid.set_table(df.head(20))
                self.Refresh()
        else:
            wx.MessageDialog(self, 'No File Path Found!',
                             'Please enter/select the file path.',
                             wx.OK|wx.CENTER).ShowModal()

    def OnLoadCSV(self, e):
        filepath = browse('CSV')
        if filepath:
            try:
                df = pandas.read_csv(filepath)
            except FileNotFoundError:
                wx.MessageDialog(self, 'Cannot download the data file.\nPlease check the file path again.',
                                 'File Not Found!', wx.OK|wx.CENTER).ShowModal()
            else:
                print(df.head())
        else:
            wx.MessageDialog(self, 'No File Path Found!',
                             'Please enter/select the file path.',
                             wx.OK|wx.CENTER).ShowModal()
