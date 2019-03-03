import wx
import os
import pandas
from components.datatable import DataGrid


class FieldCreateDialog(wx.Dialog):
    def __init__(self):
        super(FieldCreateDialog, self).__init__(
            None, -1, "Create new field", size=(600, 400))
        self.data_panel = wx.Panel(self)
        self.form_panel = wx.Panel(self)
        self.grid = DataGrid(self.data_panel)
        grid_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.data_panel, "Field values")
        grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
        self.data_panel.SetSizer(grid_sizer)

        form_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.form_panel, "Edit")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=20)
        self.field_name_lbl = wx.StaticText(self.form_panel, -1, "Name")
        self.field_name = wx.TextCtrl(self.form_panel, -1, "Column name")
        cancel_button = wx.Button(self.form_panel, -1, "Cancel")
        create_button = wx.Button(self.form_panel, -1, "Create")
        create_button.Bind(wx.EVT_BUTTON, self.onCreateButtonClick)
        cancel_button.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)
        form_sizer.AddMany([self.field_name_lbl, self.field_name])
        form_sizer.AddMany([cancel_button, create_button])
        form_box_sizer.Add(form_sizer, 1, flag=wx.EXPAND|wx.ALL)
        self.form_panel.SetSizer(form_box_sizer)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(self.data_panel, 1, flag=wx.EXPAND|wx.ALL)
        hbox.Add(self.form_panel, flag=wx.ALL|wx.EXPAND)
        self.SetSizer(hbox)

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCreateButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()


class OrganismFieldFormDialog(wx.Dialog):
    def __init__(self):
        super(OrganismFieldFormDialog, self).__init__(
            None, -1, "Create new field", size=(600, 400))
        self.data_panel = wx.Panel(self)
        self.form_panel = wx.Panel(self)
        self.grid = DataGrid(self.data_panel)
        self.grid_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.data_panel, "Field values")
        self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
        self.data_panel.SetSizer(self.grid_sizer)

        form_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.form_panel, "Edit")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=20)
        cancel_button = wx.Button(self.form_panel, -1, "Cancel")
        save_button = wx.Button(self.form_panel, -1, "Save")
        import_button = wx.Button(self.form_panel, -1, "Import..")
        save_button.Bind(wx.EVT_BUTTON, self.onSaveButtonClick)
        cancel_button.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)
        import_button.Bind(wx.EVT_BUTTON, self.onImportButtonClick)
        form_sizer.AddMany([cancel_button, save_button])
        form_sizer.Add(import_button)
        form_box_sizer.Add(form_sizer, 1, flag=wx.EXPAND|wx.ALL)
        self.form_panel.SetSizer(form_box_sizer)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(self.data_panel, 1, flag=wx.EXPAND|wx.ALL)
        hbox.Add(self.form_panel, flag=wx.ALL|wx.EXPAND)
        self.SetSizer(hbox)

        self.Bind(wx.EVT_CLOSE, self.onClose)

    def onClose(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onSaveButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()

    def onImportButtonClick(self, event):
        wildcard = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        with wx.FileDialog(None, "Choose a file", os.getcwd(),
                           "", wildcard, wx.FC_OPEN) as file_dlg:
            ret = file_dlg.ShowModal()
            file_dlg.Destroy()
            if ret == wx.ID_CANCEL:
                return
            else:
                if os.path.exists(file_dlg.GetPath()):
                    df = pandas.read_excel(file_dlg.GetPath(), usecols=2)
                    self.grid_sizer.Remove(0)
                    self.grid.Destroy()
                    self.grid = DataGrid(self.data_panel)
                    self.grid.set_table(df)
                    self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
                    self.grid_sizer.Layout()


# probably needs to extend from a base class instead for future reuse
# this needs some serious refactoring
class DrugRegFormDialog(wx.Dialog):
    def __init__(self):
        super(DrugRegFormDialog, self).__init__(
            None, -1, "Drug Registry", size=(850, 600))
        self.data_panel = wx.Panel(self)
        self.form_panel = wx.Panel(self)
        self.grid = DataGrid(self.data_panel)
        self.grid_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.data_panel, "Field values")
        self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
        self.data_panel.SetSizer(self.grid_sizer)

        form_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.form_panel, "Edit")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=20)
        add_button = wx.Button(self.form_panel, -1, "Add row")
        cancel_button = wx.Button(self.form_panel, -1, "Cancel")
        save_button = wx.Button(self.form_panel, -1, "Save")
        import_button = wx.Button(self.form_panel, -1, "Import..")
        save_button.Bind(wx.EVT_BUTTON, self.onSaveButtonClick)
        cancel_button.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)
        import_button.Bind(wx.EVT_BUTTON, self.onImportButtonClick)
        add_button.Bind(wx.EVT_BUTTON, self.onAddButtonClick)
        form_sizer.AddMany([add_button, import_button])
        form_sizer.AddMany([cancel_button, save_button])
        form_box_sizer.Add(form_sizer, 1, flag=wx.EXPAND|wx.ALL)
        self.form_panel.SetSizer(form_box_sizer)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(self.data_panel, 1, flag=wx.EXPAND|wx.ALL)
        hbox.Add(self.form_panel, flag=wx.ALL|wx.EXPAND)
        self.SetSizer(hbox)

        self.Bind(wx.EVT_CLOSE, self.onClose)

    def onClose(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onSaveButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()

    def onImportButtonClick(self, event):
        wildcard = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        with wx.FileDialog(None, "Choose a file", os.getcwd(),
                           "", wildcard, wx.FC_OPEN) as file_dlg:
            ret = file_dlg.ShowModal()
            file_dlg.Destroy()
            if ret == wx.ID_CANCEL:
                return
            else:
                if os.path.exists(file_dlg.GetPath()):
                    df = pandas.read_excel(file_dlg.GetPath(), usecols=3)
                    self.grid_sizer.Remove(0)
                    self.grid.Destroy()
                    self.grid = DataGrid(self.data_panel)
                    self.grid.set_table(df)
                    self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
                    self.grid_sizer.Layout()
                    self.grid.AutoSize()

    def onAddButtonClick(self, event):
        self.grid.AppendRows(1)
