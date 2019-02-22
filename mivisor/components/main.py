import wx
import pandas
import xlrd
from components.datatable import DataGrid


current_column = None

class FieldAttribute():
    def __init__(self, data_frame):
        self.data = {}
        for n, column in enumerate(data_frame.columns):
            self.data[column] = {'index': n,
                                 'name': column,
                                 'alias': column,
                                 'organism': False,
                                 'key': False,
                                 'drug': False,
                                 'type': str(data_frame[column].dtype),
                                 }

    @property
    def columns(self):
        return len(self.data)

    def values(self):
        return self.data.values()


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

        self.data_grid_panel = wx.Panel(self.panel, wx.ID_ANY)
        self.data_grid_panel.SetBackgroundColour("grey")
        self.summary_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Field Summary")
        self.field_attr_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Field Attributes")

        self.data_grid = DataGrid(self.data_grid_panel)

        self.edit_box = wx.StaticBox(self.panel, -1, 'Edit')

        self.summary_table = wx.ListCtrl(self.panel, style=wx.LC_REPORT)
        self.summary_table.InsertColumn(0, 'Field')
        self.summary_table.InsertColumn(1, 'Value')

        self.field_attr_list = wx.ListCtrl(self.panel, style=wx.LC_REPORT)
        self.field_attr_list.InsertColumn(0, 'Field name')
        self.field_attr_list.InsertColumn(1, 'Alias name')
        self.field_attr_list.InsertColumn(2, 'Type')
        self.field_attr_list.InsertColumn(3, 'Primary Key')
        self.field_attr_list.InsertColumn(4, 'Organism')
        self.field_attr_list.InsertColumn(5, 'Drug')
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onFieldAttrListItemSelected)

        self.summary_sizer.Add(self.summary_table, 1, wx.EXPAND)
        self.field_attr_sizer.Add(self.field_attr_list, 1, wx.EXPAND)

        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)

        self.hbox.Add(self.summary_sizer, 1, wx.EXPAND)
        self.hbox.Add(self.edit_box, 1, wx.EXPAND)
        self.vbox.Add(self.data_grid_panel, 1, wx.ALL|wx.EXPAND, 3)
        self.vbox.Add(self.field_attr_sizer, 1, wx.ALL|wx.EXPAND, 3)
        self.vbox.Add(self.hbox, 0, wx.ALL|wx.EXPAND, 3)
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
                self.data_grid.set_table(df)
                self.data_grid.Fit()
                self.field_attr = FieldAttribute(df)
                self.update_field_attrs()
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
            wx.MessageDialog(self, 'No File Path Found!',
                             'Please enter/select the file path.',
                             wx.OK|wx.CENTER).ShowModal()

    def reset_summary_table(self, desc):
        self.summary_table.ClearAll()
        self.summary_table.InsertColumn(0, 'Field')
        self.summary_table.InsertColumn(1, 'Value')
        for n,k in enumerate(desc.keys()):
            self.summary_table.InsertItem(n,k)
            self.summary_table.SetItem(n, 1, str(desc[k]))

    def onFieldAttrListItemSelected(self, evt):
        index = evt.GetIndex()
        col = self.data_grid.table.df.columns[index]
        desc = self.data_grid.table.df[col].describe()
        self.reset_summary_table(desc=desc)

    def update_field_attrs(self):
        self.field_attr_list.ClearAll()
        self.field_attr_list.InsertColumn(0, 'Field name')
        self.field_attr_list.InsertColumn(1, 'Alias name')
        self.field_attr_list.InsertColumn(2, 'Type')
        self.field_attr_list.InsertColumn(3, 'Primary Key')
        self.field_attr_list.InsertColumn(4, 'Organism')
        self.field_attr_list.InsertColumn(5, 'Drug')
        for c in sorted([co for co in self.field_attr.values()], key=lambda x: x['index']):
            self.field_attr_list.InsertItem(c['index'], c['name'])
            self.field_attr_list.SetItem(c['index'], 1, c['alias'])
            self.field_attr_list.SetItem(c['index'], 2, c['type'])
            self.field_attr_list.SetItem(c['index'], 3, str(c['key']))
            self.field_attr_list.SetItem(c['index'], 4, str(c['organism']))
            self.field_attr_list.SetItem(c['index'], 5, str(c['drug']))
