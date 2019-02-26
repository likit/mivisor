import wx
import pandas
import xlrd
from components.datatable import DataGrid
from components.fieldcreation import FieldCreateDialog


class FieldAttribute():
    def __init__(self, data_frame):
        self.data = {}
        self.columns = []
        for n, column in enumerate(data_frame.columns):
            self.columns.append(column)
            self.data[column] = {'name': column,
                                 'alias': column,
                                 'organism': False,
                                 'key': False,
                                 'drug': False,
                                 'date': False,
                                 'type': str(data_frame[column].dtype),
                                 'keep': True,
                                 'desc': "",
                                 }

    def values(self):
        return self.data.values()

    def get_column(self, colname):
        try:
            return self.data[colname]
        except KeyError as e:
            raise AttributeError(e)

    def iget_column(self, index):
        try:
            return self.columns[index]
        except IndexError:
            return None

    def get_col_index(self, colname):
        try:
            return self.columns.index(colname)
        except ValueError:
            return -1

    def is_col_aggregate(self, colname):
        if colname in self.data:
            if self.data[colname].get('aggregate'):
                return True
            else:
                return False
        else:
            raise KeyError


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
        df = pandas.DataFrame({'Name': ['Mivisor'],
                                    'Version': ['0.1'],
                                    'Description': ['User-friendly app for microbiological data analytics.'],
                                    'Contact': ['likit.pre@mahidol.edu']})

        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        dataMenu = wx.Menu()
        fieldMenu = wx.Menu()
        imp = wx.Menu()
        mlabItem = imp.Append(wx.ID_ANY, 'MLAB')
        csvItem = imp.Append(wx.ID_ANY, 'CSV')
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_ANY, 'I&mport', imp)
        exitItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit Application')
        createFieldItem = fieldMenu.Append(wx.ID_ANY, 'Aggregate')
        dataMenu.Append(wx.ID_ANY, 'Field', fieldMenu)
        menubar.Append(fileMenu, '&File')
        menubar.Append(dataMenu, '&Data')
        self.SetMenuBar(menubar)

        accel_tbl = wx.AcceleratorTable([
            (wx.ACCEL_CTRL, ord('M'), mlabItem.GetId()),
        ])
        self.SetAcceleratorTable(accel_tbl)

        self.Bind(wx.EVT_MENU, self.OnQuit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnLoadMLAB, mlabItem)
        self.Bind(wx.EVT_MENU, self.OnLoadCSV, csvItem)

        self.Bind(wx.EVT_MENU, self.OnCreateField, createFieldItem)

        # init panels
        self.preview_panel = wx.Panel(self, wx.ID_ANY)
        self.summary_panel = wx.Panel(self, wx.ID_ANY)
        self.attribute_panel = wx.Panel(self, wx.ID_ANY)
        self.edit_panel = wx.Panel(self, wx.ID_ANY)

        # init sizers
        self.summary_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.summary_panel, "Field Summary")
        self.field_attr_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.attribute_panel, "Field Attributes")
        edit_box_sizer = wx.StaticBoxSizer(wx.HORIZONTAL, self.edit_panel, "Edit")
        self.data_grid_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.preview_panel, "Data Preview")


        self.data_grid = DataGrid(self.preview_panel)
        self.data_grid.set_table(df)
        self.data_grid.AutoSizeColumns()
        self.data_grid_box_sizer.Add(self.data_grid, 1, flag=wx.EXPAND | wx.ALL)

        self.key_chkbox = wx.CheckBox(self.edit_panel, -1, label="Key", name="key")
        self.drug_chkbox = wx.CheckBox(self.edit_panel, -1, label="Drug", name="drug")
        self.organism_chkbox = wx.CheckBox(self.edit_panel, -1, label="Organism", name="organism")
        self.keep_chkbox = wx.CheckBox(self.edit_panel, -1, label="Kept", name="keep")
        self.field_edit_checkboxes = [self.key_chkbox, self.drug_chkbox, self.keep_chkbox, self.organism_chkbox]
        checkbox_sizer = wx.FlexGridSizer(cols=len(self.field_edit_checkboxes), hgap=4, vgap=0)
        for chkbox in self.field_edit_checkboxes:
            checkbox_sizer.Add(chkbox)
            chkbox.Bind(wx.EVT_CHECKBOX, self.on_edit_save_button_clicked)

        checkbox_label = wx.StaticText(self.edit_panel, -1, "Marked as")
        self.field_desc = wx.TextCtrl(self.edit_panel, -1, "", style=wx.TE_MULTILINE, size=(200, 100))
        self.field_alias = wx.TextCtrl(self.edit_panel, -1, "")
        self.edit_save_button = wx.Button(self.edit_panel, -1, "Update")
        self.edit_save_button.Bind(wx.EVT_BUTTON, self.on_edit_save_button_clicked)
        alias_label = wx.StaticText(self.edit_panel, -1, "Alias")
        desc_label = wx.StaticText(self.edit_panel, -1, "Description")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=2)
        form_sizer.AddMany([checkbox_label, checkbox_sizer])
        form_sizer.AddMany([desc_label, self.field_desc])
        form_sizer.AddMany([alias_label, self.field_alias])
        form_sizer.AddMany([wx.StaticText(self.edit_panel, -1, ""), self.edit_save_button])
        edit_box_sizer.Add(form_sizer, 1, flag=wx.ALIGN_LEFT)

        self.summary_table = wx.ListCtrl(self.summary_panel, style=wx.LC_REPORT)
        self.summary_table.InsertColumn(0, 'Field')
        self.summary_table.InsertColumn(1, 'Value')
        self.summary_sizer.Add(self.summary_table, 1, wx.EXPAND)

        self.field_attr_list = wx.ListCtrl(self.attribute_panel, style=wx.LC_REPORT)
        self.add_field_attr_list_column()
        self.field_attr_sizer.Add(self.field_attr_list, 1, wx.EXPAND)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onFieldAttrListItemSelected)

        self.preview_panel.SetSizer(self.data_grid_box_sizer)
        self.attribute_panel.SetSizer(self.field_attr_sizer)
        self.summary_panel.SetSizer(self.summary_sizer)
        self.edit_panel.SetSizer(edit_box_sizer)

        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)

        self.hbox.Add(self.edit_panel, 2, flag=wx.ALL | wx.EXPAND)
        self.hbox.Add(self.summary_panel, 1, flag=wx.ALL | wx.EXPAND)
        self.vbox.Add(self.preview_panel, 1, flag=wx.EXPAND)
        self.vbox.Add(self.attribute_panel, flag=wx.ALL | wx.EXPAND)
        self.vbox.Add(self.hbox, flag=wx.ALL | wx.EXPAND)
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
                                 'File Not Found!', wx.OK | wx.CENTER).ShowModal()
            else:
                if len(worksheets) > 1:
                    sel_worksheet = show_sheets(self, worksheets)
                else:
                    sel_worksheet = worksheets[0]
                df = pandas.read_excel(filepath, sheet_name=sel_worksheet)
                if not df.empty:
                    self.data_grid_box_sizer.Remove(0)
                    self.data_grid.Destroy()
                    self.data_grid = DataGrid(self.preview_panel)
                    self.data_grid.set_table(df)
                    self.data_grid.AutoSizeColumns()
                    self.data_grid_box_sizer.Add(self.data_grid, 1, flag=wx.EXPAND|wx.ALL)
                    self.data_grid_box_sizer.Layout()  # repaint the sizer
                    self.field_attr = FieldAttribute(df)
                    self.update_field_attrs()
        else:
            wx.MessageDialog(self, 'No File Path Found!',
                             'Please enter/select the file path.',
                             wx.OK | wx.CENTER).ShowModal()

    def OnLoadCSV(self, e):
        filepath = browse('CSV')
        if filepath:
            try:
                df = pandas.read_csv(filepath)
            except FileNotFoundError:
                wx.MessageDialog(self, 'Cannot download the data file.\nPlease check the file path again.',
                                 'File Not Found!', wx.OK | wx.CENTER).ShowModal()
        else:
            wx.MessageDialog(self, 'No File Path Found!',
                             'Please enter/select the file path.',
                             wx.OK | wx.CENTER).ShowModal()

    def OnCreateField(self, event):
        columns = []
        for c in self.field_attr.columns:
            col = self.field_attr.get_column(c)
            if col['keep']:
                columns.append(col['alias'])

        dlg = wx.SingleChoiceDialog(None,
                                    "Select a column", "Kept columns", columns)
        if dlg.ShowModal() == wx.ID_OK:
            sel_col = dlg.GetStringSelection()
        dlg.Destroy()
        if sel_col:
            sel_col_index = self.field_attr.get_col_index(sel_col)

            values = self.data_grid.table.df[sel_col].unique()
            _df = pandas.DataFrame({'Value': values, 'Group': values})
            fc = FieldCreateDialog()
            fc.grid.set_table(_df)
            resp = fc.ShowModal()

            if resp == wx.ID_OK:
                _agg_dict = {}
                for idx, row in fc.grid.table.df.iterrows():
                    _agg_dict[row['Value']] = row['Group']

                _agg_data = []
                for value in self.data_grid.table.df[sel_col]:
                    _agg_data.append(_agg_dict[value])
                new_col = fc.field_name.GetValue()
                self.data_grid.table.df.insert(sel_col_index+1, new_col, value=_agg_data)
                self.field_attr.columns.insert(sel_col_index+1, new_col)
                self.field_attr.data[new_col] = {
                    'name': new_col,
                    'alias': new_col,
                    'organism': False,
                    'key': False,
                    'drug': False,
                    'date': False,
                    'type': str(self.data_grid.table.df[new_col].dtype),
                    'keep': True,
                    'desc': "",
                    'aggregate': {
                        'from': sel_col,
                    }
                }
                self.update_field_attrs()

    def reset_summary_table(self, desc):
        self.summary_table.ClearAll()
        self.summary_table.InsertColumn(0, 'Field')
        self.summary_table.InsertColumn(1, 'Value')
        for n, k in enumerate(desc.keys()):
            self.summary_table.InsertItem(n, k)
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
        for n, c in enumerate(self.field_attr.columns):
            col = self.field_attr.get_column(c)
            self.field_attr_list.InsertItem(n, col['name'])
            self.field_attr_list.SetItem(n, 1, col['alias'])
            self.field_attr_list.SetItem(n, 2, col['type'])
            self.field_attr_list.SetItem(n, 3, str(col['key']))
            self.field_attr_list.SetItem(n, 4, str(col['date']))
            self.field_attr_list.SetItem(n, 5, str(col['organism']))
            self.field_attr_list.SetItem(n, 6, str(col['drug']))
            self.field_attr_list.SetItem(n, 7, str(col['desc']))
            self.field_attr_list.SetItem(n, 8, str(col['keep']))

    def on_edit_save_button_clicked(self, event):
        for cb in self.field_edit_checkboxes:
            name = cb.GetName()
            self.field_attr.get_column(self.current_column)[name] = cb.GetValue()
        self.field_attr.get_column(self.current_column)['alias'] = self.field_alias.GetValue()
        self.field_attr.get_column(self.current_column)['desc'] = self.field_desc.GetValue()
        self.update_field_attrs()
