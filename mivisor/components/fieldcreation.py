import wx
import os
import datetime
import pandas
from components.datatable import DataGrid
from wx.adv import DatePickerCtrl

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
        filepath = None
        wildcard = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        with wx.FileDialog(None, "Choose a file", os.getcwd(),
                           "", wildcard, wx.FC_OPEN) as file_dlg:
            ret = file_dlg.ShowModal()
            filepath = file_dlg.GetPath()

        if ret == wx.ID_CANCEL:
            return

        if filepath:
            df = pandas.read_excel(filepath, usecols=[0,1,2])
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
        filepath = None
        wildcard = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        with wx.FileDialog(None, "Choose a file", os.getcwd(),
                           "", wildcard, wx.FC_OPEN) as file_dlg:
            ret = file_dlg.ShowModal()
            filepath = file_dlg.GetPath()

        if ret == wx.ID_CANCEL:
            return

        if filepath:
            df = pandas.read_excel(filepath, usecols=[0,1,2,3])
            self.grid_sizer.Remove(0)
            self.grid.Destroy()
            self.grid = DataGrid(self.data_panel)
            self.grid.set_table(df)
            self.grid_sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL)
            self.grid_sizer.Layout()
            self.grid.AutoSize()

    def onAddButtonClick(self, event):
        self.grid.AppendRows(1)


class IndexFieldList(wx.Dialog):
    def __init__(self, choices):
        super(IndexFieldList, self).__init__(None, -1, "Antibiogram Indexes",
                                             size=(350,800))
        panel = wx.Panel(self)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        self.chlbox = wx.CheckListBox(panel, choices=choices)
        self.chlbox.Bind(wx.EVT_CHECKLISTBOX, self.onChecklistboxChecked)
        self.startDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.endDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.startDatePicker.Enable(False)
        self.endDatePicker.Enable(False)
        self.all = wx.CheckBox(panel, id=wx.ID_ANY, label="Select all dates")
        self.all.SetValue(True)

        self.rawDataIncluded = wx.CheckBox(panel, id=wx.ID_ANY, label="Include raw data in the output")
        self.rawDataIncluded.SetValue(False)

        self.indexes = []
        self.choices = choices

        self.all.Bind(wx.EVT_CHECKBOX, self.onCheckboxChecked)

        self.index_items_list = wx.ListCtrl(panel, wx.ID_ANY, style=wx.LC_LIST)

        self.ncutoff = wx.SpinCtrl(panel, value="", min=0, initial=30, name="No.Cutoff")

        label = wx.StaticText(panel, label="Select indexes:")

        vsizer.Add(label, 0, wx.ALL, 5)
        vsizer.Add(self.chlbox, 1, wx.EXPAND | wx.ALL, 5)

        nextButton = wx.Button(panel, wx.ID_OK, label="Next")
        nextButton.SetFocus()
        cancelButton = wx.Button(panel, wx.ID_CANCEL, label="Cancel")
        nextButton.Bind(wx.EVT_BUTTON, self.onNextButtonClick)
        cancelButton.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)

        staticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel, "Select a date range:")
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(cancelButton, 0, wx.EXPAND | wx.ALL, 2)
        hbox.Add(nextButton, 0, wx.EXPAND | wx.ALL, 2)

        gridbox = wx.GridSizer(2,2,2,2)

        staticBoxSizer.Add(self.all, 0, wx.EXPAND | wx.ALL, 5)
        staticBoxSizer.Add(gridbox, 0, wx.EXPAND | wx.ALL, 5)

        cutOffStaticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel,
                                                 "Set the minimum the number of isolates:")
        cutOffStaticBoxSizer.Add(self.ncutoff, wx.ALL | wx.ALIGN_LEFT, 5)

        rawDataStaticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel, "Raw data:")
        rawDataStaticBoxSizer.Add(self.rawDataIncluded, wx.ALL | wx.ALIGN_LEFT | wx.EXPAND, 5)

        startDateLabel = wx.StaticText(panel, label="Select start date")
        endDateLabel = wx.StaticText(panel, label="Select end date")
        gridbox.AddMany([startDateLabel, self.startDatePicker])
        gridbox.AddMany([endDateLabel, self.endDatePicker])
        vsizer.Add(self.index_items_list, 1, wx.EXPAND | wx.ALL, 5)
        vsizer.Add(cutOffStaticBoxSizer, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5)
        vsizer.Add(rawDataStaticBoxSizer, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5)
        vsizer.Add(staticBoxSizer, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        vsizer.Add(hbox, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(vsizer)

    def onNextButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCheckboxChecked(self, event):
        if event.IsChecked():
            self.startDatePicker.Enable(False)
            self.endDatePicker.Enable(False)
        else:
            self.startDatePicker.Enable(True)
            self.endDatePicker.Enable(True)

    def onChecklistboxChecked(self, event):
        item = event.GetInt()
        if not self.chlbox.IsChecked(item):
            idx = self.indexes.index(item)
            self.index_items_list.DeleteItem(idx)
            self.indexes.remove(item)
        else:
            self.indexes.append(item)
            self.index_items_list.Append([self.choices[item]])


class DateRangeFieldList(wx.Dialog):
    def __init__(self, parent):
        super(DateRangeFieldList, self).__init__(parent, -1, "Date Range", size=(350,250))
        panel = wx.Panel(self)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        self.startDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.endDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.startDatePicker.Enable(False)
        self.endDatePicker.Enable(False)
        self.all = wx.CheckBox(panel, id=wx.ID_ANY, label="Select all dates")
        self.all.SetValue(True)
        self.deduplicate = wx.CheckBox(panel, id=wx.ID_ANY, label="Deduplicate using a key")
        self.deduplicate.SetValue(True)

        # TODO: refactor the method's name
        self.all.Bind(wx.EVT_CHECKBOX, self.onCheckboxChecked)

        nextButton = wx.Button(panel, wx.ID_OK, label="Next")
        cancelButton = wx.Button(panel, wx.ID_CANCEL, label="Cancel")
        nextButton.Bind(wx.EVT_BUTTON, self.onNextButtonClick)
        cancelButton.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)

        staticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel, "Select a date range:")
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(cancelButton, 0, wx.EXPAND | wx.ALL, 2)
        hbox.Add(nextButton, 0, wx.EXPAND | wx.ALL, 2)

        gridbox = wx.GridSizer(2,2,2,2)

        staticBoxSizer.Add(self.deduplicate, 0, wx.EXPAND | wx.ALL, 5)
        staticBoxSizer.Add(self.all, 0, wx.EXPAND | wx.ALL, 5)
        staticBoxSizer.Add(gridbox, 0, wx.EXPAND | wx.ALL, 5)

        startDateLabel = wx.StaticText(panel, label="Select start date")
        endDateLabel = wx.StaticText(panel, label="Select end date")
        gridbox.AddMany([startDateLabel, self.startDatePicker])
        gridbox.AddMany([endDateLabel, self.endDatePicker])
        vsizer.Add(staticBoxSizer, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        vsizer.Add(hbox, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(vsizer)

    def onNextButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCheckboxChecked(self, event):
        if event.IsChecked():
            self.startDatePicker.Enable(False)
            self.endDatePicker.Enable(False)
        else:
            self.startDatePicker.Enable(True)
            self.endDatePicker.Enable(True)


class HeatmapFieldList(wx.Dialog):
    def __init__(self, choices):
        super(HeatmapFieldList, self).__init__(None, -1, "Antibiogram Heat Map",
                                             size=(350,800))
        panel = wx.Panel(self)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        self.chlbox = wx.CheckListBox(panel, choices=choices)
        self.chlbox.Bind(wx.EVT_CHECKLISTBOX, self.onChecklistboxChecked)
        self.startDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.endDatePicker = DatePickerCtrl(panel, id=wx.ID_ANY, dt=datetime.datetime.now())
        self.startDatePicker.Enable(False)
        self.endDatePicker.Enable(False)
        self.all = wx.CheckBox(panel, id=wx.ID_ANY, label="Select all dates")
        self.all.SetValue(True)

        self.indexes = []
        self.choices = choices

        self.all.Bind(wx.EVT_CHECKBOX, self.onCheckboxChecked)

        self.ncutoff = wx.SpinCtrl(panel, value="", min=0, initial=30, name="No.Cutoff")

        label = wx.StaticText(panel, label="Group by:")

        vsizer.Add(label, 0, wx.ALL, 5)
        vsizer.Add(self.chlbox, 1, wx.EXPAND | wx.ALL, 5)

        nextButton = wx.Button(panel, wx.ID_OK, label="Next")
        nextButton.SetFocus()
        cancelButton = wx.Button(panel, wx.ID_CANCEL, label="Cancel")
        nextButton.Bind(wx.EVT_BUTTON, self.onNextButtonClick)
        cancelButton.Bind(wx.EVT_BUTTON, self.onCancelButtonClick)

        staticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel, "Select a date range:")
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(cancelButton, 0, wx.EXPAND | wx.ALL, 2)
        hbox.Add(nextButton, 0, wx.EXPAND | wx.ALL, 2)

        gridbox = wx.GridSizer(2,2,2,2)

        staticBoxSizer.Add(self.all, 0, wx.EXPAND | wx.ALL, 5)
        staticBoxSizer.Add(gridbox, 0, wx.EXPAND | wx.ALL, 5)

        cutOffStaticBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, panel,
                                                 "Set the minimum the number of isolates:")
        cutOffStaticBoxSizer.Add(self.ncutoff, wx.ALL | wx.ALIGN_LEFT, 5)

        startDateLabel = wx.StaticText(panel, label="Select start date")
        endDateLabel = wx.StaticText(panel, label="Select end date")
        gridbox.AddMany([startDateLabel, self.startDatePicker])
        gridbox.AddMany([endDateLabel, self.endDatePicker])
        vsizer.Add(cutOffStaticBoxSizer, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5)
        vsizer.Add(staticBoxSizer, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        vsizer.Add(hbox, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(vsizer)

    def onNextButtonClick(self, event):
        self.EndModal(wx.ID_OK)
        self.Destroy()

    def onCancelButtonClick(self, event):
        self.EndModal(wx.ID_CANCEL)
        self.Destroy()

    def onCheckboxChecked(self, event):
        if event.IsChecked():
            self.startDatePicker.Enable(False)
            self.endDatePicker.Enable(False)
        else:
            self.startDatePicker.Enable(True)
            self.endDatePicker.Enable(True)

    def onChecklistboxChecked(self, event):
        item = event.GetInt()
        if not self.chlbox.IsChecked(item):
            self.indexes.remove(item)
        else:
            self.indexes.append(item)
            for i in self.chlbox.GetCheckedItems():
                if i != item:
                    self.chlbox.Check(i, check=False)
                    self.indexes.remove(i)
