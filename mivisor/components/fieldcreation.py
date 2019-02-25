import wx
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

