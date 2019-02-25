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
                                 'date': False,
                                 'type': str(data_frame[column].dtype),
                                 'keep': True,
                                 'desc': ""
                                 }

    @property
    def columns(self):
        return len(self.data)

    def values(self):
        return self.data.values()

    def get_column(self, colname):
        try:
            return self.data[colname]
        except KeyError as e:
            raise AttributeError(e)


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

        self.current_column = None

        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        imp = wx.Menu()
        mlabItem = imp.Append(wx.ID_ANY, 'MLAB')
        csvItem = imp.Append(wx.ID_ANY, 'CSV')
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_ANY, 'I&mport', imp)
        exitItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit Application')
        menubar.Append(fileMenu, '&File')
        self.SetMenuBar(menubar)

        accel_tbl = wx.AcceleratorTable([
            (wx.ACCEL_CTRL, ord('M'), mlabItem.GetId()),
        ])
        self.SetAcceleratorTable(accel_tbl)

        self.Bind(wx.EVT_MENU, self.OnQuit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnLoadMLAB, mlabItem)
        self.Bind(wx.EVT_MENU, self.OnLoadCSV, csvItem)

        # init panels
        self.preview_panel = wx.Panel(self, wx.ID_ANY)
        self.summary_panel = wx.Panel(self, wx.ID_ANY)
        self.attribute_panel = wx.Panel(self, wx.ID_ANY)
        self.edit_panel = wx.Panel(self, wx.ID_ANY)

        # init sizers
        self.summary_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.summary_panel, "Field Summary")
        self.field_attr_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.attribute_panel, "Field Attributes")
        self.edit_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.edit_panel, "Edit")

        self.summary_panel.SetSizer(self.summary_sizer)
        self.attribute_panel.SetSizer(self.field_attr_sizer)
        self.edit_panel.SetSizer(self.edit_box_sizer)

        self.data_grid_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.preview_panel, "Data Preview")
        self.data_grid = DataGrid(self.preview_panel)
        self.data_grid_box_sizer.Add(self.data_grid, 1, flag=wx.EXPAND|wx.ALL)
        self.preview_panel.SetSizer(self.data_grid_box_sizer)

        self.key_chkbox = wx.CheckBox(self.edit_panel, -1, label="Key", name="key")
        self.drug_chkbox = wx.CheckBox(self.edit_panel, -1, label="Drug", name="drug")
        self.organism_chkbox = wx.CheckBox(self.edit_panel, -1, label="Organism", name="organism")
        self.keep_chkbox = wx.CheckBox(self.edit_panel, -1, label="Kept", name="keep")
        self.field_edit_checkboxes = [self.key_chkbox, self.drug_chkbox, self.keep_chkbox, self.organism_chkbox]
        checkbox_sizer = wx.FlexGridSizer(cols=len(self.field_edit_checkboxes), hgap=4, vgap=0)
        for chkbox in self.field_edit_checkboxes:
            checkbox_sizer.Add(chkbox)
            chkbox.Bind(wx.EVT_CHECKBOX, self.on_edit_save_button_clicked)

        self.field_desc = wx.TextCtrl(self.edit_panel, -1, "", style=wx.TE_MULTILINE, size=(200,100))
        self.field_alias = wx.TextCtrl(self.edit_panel, -1, "")
        edit_save_button = wx.Button(self.edit_panel, -1, "Update")
        edit_save_button.Bind(wx.EVT_BUTTON, self.on_edit_save_button_clicked)

        alias_label = wx.StaticText(self.edit_panel, -1, "Alias")
        desc_label = wx.StaticText(self.edit_panel, -1, "Description")
        checkbox_label = wx.StaticText(self.edit_panel, -1, "Marked as")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=2)
        form_sizer.AddMany([checkbox_label, checkbox_sizer])
        form_sizer.AddMany([desc_label, self.field_desc])
        form_sizer.AddMany([alias_label, self.field_alias])
        self.edit_box_sizer.Add(checkbox_sizer, 0, flag=wx.ALIGN_LEFT)
        self.edit_box_sizer.Add(form_sizer, 0, flag=wx.ALIGN_LEFT)
        self.edit_box_sizer.Add(edit_save_button, 0, flag=wx.ALIGN_CENTER)


        self.summary_table = wx.ListCtrl(self.summary_panel, style=wx.LC_REPORT)
        self.summary_table.InsertColumn(0, 'Field')
        self.summary_table.InsertColumn(1, 'Value')

        self.field_attr_list = wx.ListCtrl(self.attribute_panel, style=wx.LC_REPORT)
        self.add_field_attr_list_column()
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onFieldAttrListItemSelected)

        self.summary_sizer.Add(self.summary_table, 1, wx.EXPAND)
        self.field_attr_sizer.Add(self.field_attr_list, 1, wx.EXPAND)

        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)

        self.hbox.Add(self.summary_panel, 1, flag=wx.EXPAND)
        self.hbox.Add(self.edit_panel, 1, flag=wx.EXPAND)
        self.vbox.Add(self.preview_panel, 1, flag=wx.EXPAND)
        self.vbox.Add(self.attribute_panel, flag=wx.EXPAND)
        self.vbox.Add(self.hbox, flag=wx.ALL|wx.EXPAND)
        self.SetSizer(self.vbox)


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
                # self.data_grid.Fit()
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
        self.current_column = self.data_grid.table.df.columns[index]
        desc = self.data_grid.table.df[self.current_column].describe()
        self.reset_summary_table(desc=desc)
        for cb in self.field_edit_checkboxes:
            name = cb.GetName()
            cb.SetValue(self.field_attr.get_column(self.current_column)[name])
        self.field_alias.SetValue(self.field_attr.get_column(self.current_column)['alias'])
        self.field_desc.SetValue(self.field_attr.get_column(self.current_column)['desc'])
        self.data_grid.SelectCol(index)

    def add_field_attr_list_column(self):
        self.field_attr_list.ClearAll()
        self.field_attr_list.InsertColumn(0, 'Field name')
        self.field_attr_list.InsertColumn(1, 'Alias name')
        self.field_attr_list.InsertColumn(2, 'Type')
        self.field_attr_list.InsertColumn(3, 'Key')
        self.field_attr_list.InsertColumn(4, 'Date')
        self.field_attr_list.InsertColumn(5, 'Organism')
        self.field_attr_list.InsertColumn(6, 'Drug')
        self.field_attr_list.InsertColumn(7, 'Description')
        self.field_attr_list.InsertColumn(8, 'Keep')
        self.field_attr_list.SetColumnWidth(7, 300)

    def update_field_attrs(self):
        for c in sorted([co for co in self.field_attr.values()], key=lambda x: x['index']):
            self.field_attr_list.InsertItem(c['index'], c['name'])
            self.field_attr_list.SetItem(c['index'], 1, c['alias'])
            self.field_attr_list.SetItem(c['index'], 2, c['type'])
            self.field_attr_list.SetItem(c['index'], 3, str(c['key']))
            self.field_attr_list.SetItem(c['index'], 4, str(c['date']))
            self.field_attr_list.SetItem(c['index'], 5, str(c['organism']))
            self.field_attr_list.SetItem(c['index'], 6, str(c['drug']))
            self.field_attr_list.SetItem(c['index'], 7, str(c['desc']))
            self.field_attr_list.SetItem(c['index'], 8, str(c['keep']))

    def on_edit_save_button_clicked(self, event):
        for cb in self.field_edit_checkboxes:
            name = cb.GetName()
            self.field_attr.get_column(self.current_column)[name] = cb.GetValue()
        self.field_attr.get_column(self.current_column)['alias'] = self.field_alias.GetValue()
        self.field_attr.get_column(self.current_column)['desc'] = self.field_desc.GetValue()
        self.update_field_attrs()
