import wx
import pandas as pd
import os
from .datatable import DataGrid


class DrugRegFormDialog(wx.Dialog):
    def __init__(self):
        super(DrugRegFormDialog, self).__init__(None, -1, "Drug Registry", size=(850, 600))
        self.data_panel = wx.Panel(self)
        self.form_panel = wx.Panel(self)
        self.grid = DataGrid(self.data_panel)
        self.grid_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.data_panel, "Field values")
        self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND | wx.ALL)
        self.data_panel.SetSizer(self.grid_sizer)
        self.load_drug_registry()

        form_box_sizer = wx.StaticBoxSizer(wx.VERTICAL, self.form_panel, "Edit")
        form_sizer = wx.FlexGridSizer(cols=2, hgap=2, vgap=20)
        add_button = wx.Button(self.form_panel, -1, "Add")
        delete_button = wx.Button(self.form_panel, -1, "Delete")
        cancel_button = wx.Button(self.form_panel, -1, "Cancel")
        save_button = wx.Button(self.form_panel, -1, "Save")
        save_button.Bind(wx.EVT_BUTTON, self.onSaveButtonClick)
        cancel_button.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)
        add_button.Bind(wx.EVT_BUTTON, self.onAddButtonClick)
        delete_button.Bind(wx.EVT_BUTTON, self.onDeleteButtonClick)
        form_sizer.AddMany([add_button, delete_button, cancel_button, save_button])
        form_box_sizer.Add(form_sizer, 1, flag=wx.EXPAND | wx.ALL)
        self.form_panel.SetSizer(form_box_sizer)
        form_box_sizer.Fit(self.form_panel)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(self.data_panel, 1, flag=wx.EXPAND | wx.ALL)
        hbox.Add(self.form_panel, flag=wx.ALL | wx.EXPAND)
        self.SetSizer(hbox)

        self.Bind(wx.EVT_CLOSE, self.onClose)

    def onClose(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onSaveButtonClick(self, event):
        try:
            self.grid.table.df.to_json(os.path.join('appdata', 'drugs.json'))
        except:
            pass
        else:
            if wx.MessageBox('Drugs data saved.', style=wx.OK) == wx.OK:
                self.update_drug_list()
                self.EndModal(wx.ID_OK)
                self.Destroy()

    def onAddButtonClick(self, event):
        self.grid.AppendRows(1)

    def onDeleteButtonClick(self, event):
        row_idx = self.grid.GetGridCursorRow()
        with wx.MessageDialog(self,
                              'Do you want to delete {}?'.format(self.grid.table.df.iloc[row_idx]['drug']),
                              style=wx.YES_NO) as dlg:
            if dlg.ShowModal() == wx.ID_YES:
                self.grid.DeleteRows(row_idx)

    def load_drug_registry(self):
        try:
            self.drug_df = pd.read_json(os.path.join('appdata', 'drugs.json'))
            self.grid.set_table(self.drug_df)
            self.grid.AutoSize()
        except:
            return pd.DataFrame(columns=['drug', 'abbreviation', 'group'])
        else:
            if self.drug_df.empty:
                self.drug_df = pd.DataFrame(columns=['drug', 'abbreviation', 'group'])
            self.update_drug_list()

    def update_drug_list(self):
        drug_list = []
        self.drug_df = self.drug_df.sort_values(['group'])
        for idx, row in self.drug_df.iterrows():
            if row['abbreviation']:
                abbrs = [a.strip().upper() for a in row['abbreviation'].split(',')]
            else:
                abbrs = []
            for ab in abbrs:
                drug_list.append({'drug': row['drug'], 'abbr': ab, 'group': row['group']})
        self.drug_data = pd.DataFrame(drug_list)
