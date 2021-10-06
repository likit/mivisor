import os
import sys

import wx
import wx.adv
import pandas as pd
import xlsxwriter
from ObjectListView import ObjectListView, ColumnDefn, FastObjectListView
from threading import Thread
from pubsub import pub

from components.drug_dialog import DrugRegFormDialog


CLOSE_PROGRESS_BAR_SIGNAL = 'close-progressbar'
WRITE_TO_EXCEL_FILE_SIGNAL = 'write-to-excel-file'
ENABLE_BUTTONS = 'enable-buttons'
DISABLE_BUTTONS = 'disable-buttons'

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class PulseProgressBarDialog(wx.ProgressDialog):
    def __init__(self, *args, abort_message='abort'):
        super(PulseProgressBarDialog, self)\
            .__init__(*args, style=wx.PD_AUTO_HIDE | wx.PD_APP_MODAL)
        pub.subscribe(self.close, CLOSE_PROGRESS_BAR_SIGNAL)
        while self.GetValue() != self.GetRange():
            self.Pulse()

    def close(self):
        self.Update(self.GetRange())


class ReadExcelThread(Thread):
    def __init__(self, filepath, message):
        super(ReadExcelThread, self).__init__()
        self._filepath = filepath
        self._message = message
        self.start()

    def run(self):
        df = pd.read_excel(self._filepath)
        df = df.dropna(how='all').fillna('')
        wx.CallAfter(pub.sendMessage, self._message, df=df)


class BiogramGeneratorThread(Thread):
    def __init__(self, data, date_col, identifier_col, organism_col, indexes, keys,
                 include_count, include_percent, include_narst, columns, drug_data):
        super(BiogramGeneratorThread, self).__init__()
        self.drug_data = drug_data
        self.data = data
        self.date_col = date_col
        self.identifier_col = identifier_col
        self.organism_col = organism_col
        self.columns = columns
        self.indexes = indexes
        self.keys = keys
        self.include_count = include_count
        self.include_percent = include_percent
        self.include_narst = include_narst
        self.start()

    def run(self):
        # TODO: remove hard-coded organism file
        organism_df = pd.read_excel(os.path.join('appdata', 'organisms2020.xlsx'))

        melted_df = self.data.melt(id_vars=self.keys)
        _melted_df = pd.merge(melted_df, organism_df, how='inner')
        _melted_df = pd.merge(_melted_df, self.drug_data,
                              right_on='abbr', left_on='variable', how='outer')
        indexes = [self.columns[idx] for idx in self.indexes]
        total = _melted_df.pivot_table(index=indexes,
                                       columns=['group', 'variable'], aggfunc='count')
        sens = _melted_df[_melted_df['value'] == 'S'].pivot_table(index=indexes,
                                                                  columns=['group', 'variable'],
                                                                  aggfunc='count')
        resists = _melted_df[(_melted_df['value'] == 'I') | (_melted_df['value'] == 'R')] \
            .pivot_table(index=indexes, columns=['group', 'variable'], aggfunc='count')
        biogram_resists = (resists / total * 100).applymap(lambda x: round(x, 2))
        biogram_sens = (sens / total * 100).applymap(lambda x: round(x, 2))
        formatted_total = total.applymap(lambda x: '' if pd.isna(x) else '{:.0f}'.format(x))
        biogram_narst_s = biogram_sens.fillna('-').applymap(str) + " (" + formatted_total + ")"
        biogram_narst_s = biogram_narst_s.applymap(lambda x: '' if x.startswith('-') else x)
        wx.CallAfter(pub.sendMessage, CLOSE_PROGRESS_BAR_SIGNAL)
        wx.CallAfter(pub.sendMessage,
                     WRITE_TO_EXCEL_FILE_SIGNAL,
                     sens=sens if self.include_count else None,
                     resists=resists if self.include_count else None,
                     biogram_sens=biogram_sens if self.include_percent else None,
                     biogram_resists=biogram_resists if self.include_percent else None,
                     biogram_narst_s=biogram_narst_s if self.include_narst else None,
                     identifier_col=self.identifier_col)


class DataRow(object):
    def __init__(self, id, series) -> None:
        self.id = id
        for k, v in series.items():
            setattr(self, k, v)

    def to_list(self, columns):
        return [getattr(self, c) for c in columns]

    def to_dict(self, columns):
        return {c: getattr(self, c) for c in columns}


class BiogramIndexDialog(wx.Dialog):
    def __init__(self, parent, columns, title='Biogram Indexes', start=None, end=None):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.indexes = []
        self.choices = columns
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        instruction = wx.StaticText(self, label='Data will be organized in hierarchy based on the order in the list.')
        self.chlbox = wx.CheckListBox(self, choices=columns)
        self.chlbox.Bind(wx.EVT_CHECKLISTBOX, self.on_checked)
        self.index_items_list = wx.ListCtrl(self, wx.ID_ANY, style=wx.LC_REPORT, size=(300, 200))
        self.index_items_list.AppendColumn('Level')
        self.index_items_list.AppendColumn('Attribute')
        dateBoxSizer = wx.StaticBoxSizer(wx.HORIZONTAL, self, label='Date Range')
        self.startDate = wx.adv.DatePickerCtrl(self, dt=start)
        startDateLabel = wx.StaticText(self, label='Start')
        self.endDate = wx.adv.DatePickerCtrl(self, dt=end)
        endDateLabel = wx.StaticText(self, label='End')
        dateBoxSizer.Add(startDateLabel, 0, wx.ALL, 5)
        dateBoxSizer.Add(self.startDate, 0, wx.ALL, 5)
        dateBoxSizer.Add(endDateLabel, 0, wx.ALL, 5)
        dateBoxSizer.Add(self.endDate, 0, wx.ALL, 5)
        outputBoxSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label='Output')
        self.includeCount = wx.CheckBox(self, label='Raw counts')
        self.includePercent = wx.CheckBox(self, label='Percents')
        self.includeNarstStyle = wx.CheckBox(self, label='NARST format')
        self.includeCount.SetValue(True)
        self.includePercent.SetValue(True)
        self.includeNarstStyle.SetValue(True)
        outputBoxSizer.Add(self.includeCount, 0, wx.ALL, 5)
        outputBoxSizer.Add(self.includePercent, 0, wx.ALL, 5)
        outputBoxSizer.Add(self.includeNarstStyle, 0, wx.ALL, 5)
        main_sizer.Add(instruction, 0, wx.ALL, 5)
        main_sizer.Add(self.chlbox, 1, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.index_items_list, 1, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(dateBoxSizer, 0, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(outputBoxSizer, 0, wx.ALL | wx.EXPAND, 5)
        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(self, id=wx.ID_OK, label='Generate')
        ok_btn.SetDefault()
        cancel_btn = wx.Button(self, id=wx.ID_CANCEL, label='Cancel')
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        main_sizer.SetSizeHints(self)
        self.SetSizer(main_sizer)
        main_sizer.Fit(self)

    def on_checked(self, event):
        item = event.GetInt()
        if not self.chlbox.IsChecked(item):
            idx = self.indexes.index(item)
            self.index_items_list.DeleteItem(idx)
            self.indexes.remove(item)
        else:
            self.indexes.append(item)
            self.index_items_list.Append([len(self.indexes), self.choices[item]])


class NewColumnDialog(wx.Dialog):
    def __init__(self, parent, data, title='Edit values and save to a new column'):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        colname_label = wx.StaticText(self, label='New column name')
        self.colname_ctrl = wx.TextCtrl(self)
        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(self, id=wx.ID_OK, label='Create')
        ok_btn.SetDefault()
        cancel_btn = wx.Button(self, id=wx.ID_CANCEL, label='Cancel')
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        self.olvData = ObjectListView(self, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        self.olvData.oddRowsBackColor = wx.Colour(230, 230, 230, 100)
        self.olvData.evenRowsBackColor = wx.WHITE
        self.olvData.cellEditMode = ObjectListView.CELLEDIT_DOUBLECLICK
        self.data = []
        for dt in data:
            self.data.append({'old': dt, 'new': dt})
        self.olvData.SetColumns([
            ColumnDefn(title='Old Value', align='left', minimumWidth=50, valueGetter='old'),
            ColumnDefn(title='New Value', align='left', minimumWidth=50, valueGetter='new'),
        ])
        self.olvData.SetObjects(self.data)
        main_sizer.Add(self.olvData, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(colname_label, 0, wx.ALL, 5)
        main_sizer.Add(self.colname_ctrl, 0, wx.ALL, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        self.SetAutoLayout(True)
        self.SetSizer(main_sizer)
        main_sizer.Fit(self)

    def replace(self):
        return self.data


class DrugListCtrl(wx.ListCtrl):
    def __init__(self, parent, cols):
        super(DrugListCtrl, self).__init__(parent, style=wx.LC_REPORT, size=(300, 200))
        self.EnableCheckBoxes(True)
        self.Bind(wx.EVT_LIST_ITEM_CHECKED, self.on_check)
        self.Bind(wx.EVT_LIST_ITEM_UNCHECKED, self.on_uncheck)
        self.cols = cols
        self.drugs = []
        self.AppendColumn('Name')
        for col in cols:
            self.Append([col])

        for d in config.Read('Drugs').split(';'):
            if d in self.cols:
                idx = self.cols.index(d)
                self.CheckItem(idx)

    def on_check(self, event):
        item = event.GetItem()
        idx = item.GetId()
        col = self.cols[idx]
        if self.IsItemChecked(idx):
            self.drugs.append(col)

    def on_uncheck(self, event):
        item = event.GetItem()
        idx = item.GetId()
        col = self.cols[idx]
        if col in self.drugs:
            self.drugs.remove(col)


class ConfigDialog(wx.Dialog):
    def __init__(self, parent, columns, title='Configuration'):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        form_sizer = wx.FlexGridSizer(5, 2, 15, 20)
        form_sizer.Add(wx.StaticText(self, id=wx.ID_ANY, label='Identifier'), 0)
        self.identifier_combo_ctrl = wx.Choice(self, choices=columns)
        _col = config.Read('IdentifierCol', '')
        if _col and _col in columns:
            self.identifier_combo_ctrl.SetSelection(columns.index(_col))
        form_sizer.Add(self.identifier_combo_ctrl, 0)
        form_sizer.Add(wx.StaticText(self, id=wx.ID_ANY, label='Date'))
        self.date_combo_ctrl = wx.Choice(self, choices=columns)
        _col = config.Read('DateCol', '')
        if _col and _col in columns:
            self.date_combo_ctrl.SetSelection(columns.index(_col))
        form_sizer.Add(self.date_combo_ctrl, 0)
        form_sizer.Add(wx.StaticText(self, id=wx.ID_ANY, label='Organism Code'))
        self.organism_combo_ctrl = wx.Choice(self, choices=columns)
        _col = config.Read('OrganismCol', '')
        if _col and _col in columns:
            self.organism_combo_ctrl.SetSelection(columns.index(_col))
        form_sizer.Add(self.organism_combo_ctrl, 0)
        form_sizer.Add(wx.StaticText(self, id=wx.ID_ANY, label='Specimens'))
        self.specimens_combo_ctrl = wx.Choice(self, choices=columns)
        _col = config.Read('SpecimensCol', '')
        if _col and _col in columns:
            self.specimens_combo_ctrl.SetSelection(columns.index(_col))
        form_sizer.Add(self.specimens_combo_ctrl, 0)
        self.drug_listctrl = DrugListCtrl(self, columns)
        form_sizer.Add(wx.StaticText(self, id=wx.ID_ANY, label='Drugs'))
        form_sizer.Add(self.drug_listctrl, 1, wx.EXPAND)

        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(self, id=wx.ID_OK, label='Ok')
        ok_btn.SetDefault()
        cancel_btn = wx.Button(self, id=wx.ID_CANCEL, label='Cancel')
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()

        main_sizer.Add(form_sizer, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        main_sizer.SetSizeHints(self)
        self.SetSizer(main_sizer)
        main_sizer.Fit(self)


class DeduplicateIndexDialog(wx.Dialog):
    def __init__(self, parent, columns, title='Deduplication Keys'):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.keys = []
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        instruction = wx.StaticText(self, label='Select columns you want to use for deduplication.')
        self.isSortDate = wx.CheckBox(self, label='Sort by the date column')
        self.isSortDate.SetValue(True)
        self.chlbox = wx.CheckListBox(self, choices=columns)
        self.chlbox.Bind(wx.EVT_CHECKLISTBOX, self.on_checked)
        button_sizer = self.CreateStdDialogButtonSizer(flags=wx.OK | wx.CANCEL)
        main_sizer.Add(instruction, 0, wx.ALL, 5)
        main_sizer.Add(self.chlbox, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(self.isSortDate, 0, wx.ALL, 5)
        main_sizer.Add(button_sizer, 0, wx.ALL, 5)
        self.SetSizer(main_sizer)
        self.Fit()

    def on_checked(self, event):
        item = event.GetInt()
        if not self.chlbox.IsChecked(item):
            idx = self.keys.index(item)
            self.keys.remove(item)
        else:
            self.keys.append(item)


class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY,
                          title="Mivisor Version 2021.1", size=(800, 600))
        panel = wx.Panel(self)
        # TODO: figure out how to update the statusbar's text from the frame's children
        self.statusbar = self.CreateStatusBar(2)
        self.statusbar.SetStatusText('The app is ready to roll.')
        self.statusbar.SetStatusText('This is for the analytics information', 1)
        menuBar = wx.MenuBar()
        fileMenu = wx.Menu()
        registryMenu = wx.Menu()
        menuBar.Append(fileMenu, '&File')
        menuBar.Append(registryMenu, 'Re&gistry')
        fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit Application')
        exportItem = fileMenu.Append(wx.ID_EXIT, 'Export Data', 'Export Data')
        drugItem = registryMenu.Append(wx.ID_ANY, 'Drugs', 'Drug Registry')
        self.SetMenuBar(menuBar)
        self.Bind(wx.EVT_MENU, lambda x: self.Close(), fileItem)
        self.Bind(wx.EVT_MENU, self.open_drug_dialog, drugItem)
        self.Bind(wx.EVT_MENU, self.export_data, exportItem)

        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.SetIcon(wx.Icon(resource_path(os.path.join('icons', 'appicon.ico'))))
        self.Center()
        self.Maximize(True)

        self.load_drug_data()

        self.df = pd.DataFrame()
        self.data = []
        self.colnames = []
        self.organism_col = config.Read('OrganismCol', '')
        self.identifier_col = config.Read('IdentifierCol', '')
        self.date_col = config.Read('DateCol', '')
        self.specimens_col = config.Read('SpecimensCol', '')
        self.drugs_col = config.Read('Drugs', '').split(';') or []
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        load_button = wx.Button(panel, label="Load")
        self.copy_button = wx.Button(panel, label="Copy Column")
        self.config_btn = wx.Button(panel, label='Config')
        self.generate_btn = wx.Button(panel, label='Generate')
        load_button.Bind(wx.EVT_BUTTON, self.open_load_data_dialog)
        self.copy_button.Bind(wx.EVT_BUTTON, self.copy_column)
        self.config_btn.Bind(wx.EVT_BUTTON, self.configure)
        self.generate_btn.Bind(wx.EVT_BUTTON, self.generate)

        self.dataOlv = FastObjectListView(panel, wx.ID_ANY,
                                          style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        self.dataOlv.oddRowsBackColor = wx.Colour(230, 230, 230, 100)
        self.dataOlv.evenRowsBackColor = wx.WHITE
        self.dataOlv.cellEditMode = ObjectListView.CELLEDIT_DOUBLECLICK
        self.dataOlv.SetEmptyListMsg('Welcome to Mivisor Version 2021.1')
        self.dataOlv.SetObjects([])
        main_sizer.Add(self.dataOlv, 1, wx.ALL | wx.EXPAND, 10)
        btn_sizer.Add(load_button, 0, wx.ALL, 5)
        btn_sizer.Add(self.copy_button, 0, wx.ALL, 5)
        btn_sizer.Add(self.config_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.generate_btn, 0, wx.ALL, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALL, 5)
        panel.SetSizer(main_sizer)
        panel.Fit()

        self.disable_buttons()

        pub.subscribe(self.disable_buttons, DISABLE_BUTTONS)
        pub.subscribe(self.enable_buttons, ENABLE_BUTTONS)
        pub.subscribe(self.write_output, WRITE_TO_EXCEL_FILE_SIGNAL)

    def OnClose(self, event):
        if event.CanVeto():
            if wx.MessageBox('You want to quit the program?', 'Please confirm', style=wx.YES_NO) != wx.YES:
                event.Veto()
                return
        event.Skip()

    def disable_buttons(self):
        self.generate_btn.Disable()
        self.copy_button.Disable()
        self.config_btn.Disable()

    def enable_buttons(self):
        self.generate_btn.Enable()
        self.copy_button.Enable()
        self.config_btn.Enable()

    def export_data(self, event):
        df = pd.DataFrame([d.to_dict(self.colnames) for d in self.data])
        with wx.FileDialog(self, "Please select the output file for data",
                           wildcard="Excel file (*xlsx)|*xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            file_path = file_dialog.GetPath()
            if os.path.splitext(file_path)[1] != '.xlsx':
                file_path = file_path + '.xlsx'
            try:
                df.to_excel(file_path, index=False)
            except:
                with wx.MessageDialog(self, 'Export failed.', 'Export Data', style=wx.OK) as dlg:
                    dlg.ShowModal()
            else:
                with wx.MessageDialog(self, 'Export completed.', 'Export Data', style=wx.OK) as dlg:
                    dlg.ShowModal()

    def load_drug_data(self):
        try:
            drug_df = pd.read_json(os.path.join('appdata', 'drugs.json'))
        except:
            pass
        if drug_df.empty:
            drug_df = pd.DataFrame(columns=['drug', 'abbreviation', 'group'])

        drug_list = []
        drug_df = drug_df.sort_values(['group'])
        for idx, row in drug_df.iterrows():
            if row['abbreviation']:
                abbrs = [a.strip().upper() for a in row['abbreviation'].split(',')]
            else:
                abbrs = []
            for ab in abbrs:
                drug_list.append({'drug': row['drug'], 'abbr': ab, 'group': row['group']})
        self.drug_data = pd.DataFrame(drug_list)

    def set_data_olv(self, df):
        self.df = df
        self.df = self.df.dropna(how='all').fillna('')
        self.setColumns()
        self.data = [DataRow(idx, row) for idx, row in self.df.iterrows()]
        self.dataOlv.SetObjects(self.data)
        pub.sendMessage(CLOSE_PROGRESS_BAR_SIGNAL)
        pub.sendMessage(ENABLE_BUTTONS)

    def read_data_from_file(self):
        with wx.FileDialog(self, "Load data from file",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
                           wildcard="Excel file (*.xlsx)|*.xlsx") as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            filepath = file_dialog.GetPath()
            pub.subscribe(self.set_data_olv, 'load_excel_data_finished')
            ReadExcelThread(filepath, 'load_excel_data_finished')
            progress_bar = PulseProgressBarDialog('Loading Data', 'Reading data from {}'.format(filepath))

    def open_load_data_dialog(self, event):
        pub.sendMessage(DISABLE_BUTTONS)
        if not self.df.empty:
            with wx.MessageDialog(self, "Load new dataset?", "Load data", style=wx.YES_NO) as msg_dialog:
                if msg_dialog.ShowModal() == wx.ID_YES:
                    self.read_data_from_file()
        else:
            self.read_data_from_file()

    def open_drug_dialog(self, event):
        with DrugRegFormDialog() as drug_dlg:
            drug_dlg.ShowModal()

    def setColumns(self):
        columns = []
        for c in self.df.columns:
            self.colnames.append(c)
            col_type = str(self.df.dtypes.get(c))
            if col_type.startswith('int') or col_type.startswith('float'):
                formatter = '%.1f'
            elif col_type.startswith('datetime'):
                formatter = '%Y-%m-%d'
            else:
                formatter = '%s'
            columns.append(
                ColumnDefn(title=c.title(), align='left', stringConverter=formatter, valueGetter=c, minimumWidth=50))
        self.dataOlv.SetColumns(columns)

    def copy_column(self, event):
        with wx.SingleChoiceDialog(self, 'Select Source Column', 'Source Column', choices=self.colnames) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                idx = dlg.GetSelection()
                colname = self.colnames[idx]
                data = set([getattr(row, colname) for row in self.data])
                with NewColumnDialog(self, data) as dlg:
                    if dlg.ShowModal() == wx.ID_OK:
                        new_data = dlg.replace()
                        lookup_dict = {}
                        new_colname = dlg.colname_ctrl.GetValue()
                        for item in new_data:
                            lookup_dict[item['old']] = item['new']
                        for d in self.data:
                            old_value = getattr(d, colname)
                            setattr(d, new_colname, lookup_dict.get(old_value))
                        self.dataOlv.AddColumnDefn(ColumnDefn(
                            title=new_colname,
                            align='left',
                            valueGetter=new_colname,
                            minimumWidth=50,
                        ))
                        self.colnames.append(new_colname)
                        self.dataOlv.RepopulateList()

    def configure(self, event):
        with ConfigDialog(self, self.colnames) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                if dlg.identifier_combo_ctrl.GetSelection() != -1:
                    self.identifier_col = self.colnames[dlg.identifier_combo_ctrl.GetSelection()]
                if dlg.date_combo_ctrl.GetSelection() != -1:
                    self.date_col = self.colnames[dlg.date_combo_ctrl.GetSelection()]
                if dlg.organism_combo_ctrl.GetSelection() != -1:
                    self.organism_col = self.colnames[dlg.organism_combo_ctrl.GetSelection()]
                if dlg.specimens_combo_ctrl.GetSelection() != -1:
                    self.specimens_col = self.colnames[dlg.specimens_combo_ctrl.GetSelection()]
                self.drugs_col = dlg.drug_listctrl.drugs
                config.Write('IdentifierCol', self.identifier_col)
                config.Write('DateCol', self.date_col)
                config.Write('OrganismCol', self.organism_col)
                config.Write('SpecimensCol', self.specimens_col)
                config.Write('Drugs', ';'.join(self.drugs_col))

    def melt(self, source_data=None):
        if source_data is None:
            source_data = self.data
        keys = []
        for c in self.colnames:
            if c not in self.drugs_col:
                keys.append(c)
        data = [row.to_list(self.colnames) for row in source_data]
        ids = [row.id for row in data]
        df = pd.DataFrame(data=data, index=ids, columns=self.colnames)
        return df.melt(id_vars=keys)

    def generate(self, event):
        if not all([self.date_col, self.identifier_col, self.organism_col]):
            self.configure()
        df = pd.DataFrame([d.to_dict(self.colnames) for d in self.data])
        if df.empty:
            with wx.MessageDialog(self, 'No data provided. Please load data from an Excel file',
                                  'Error', style=wx.OK) as dlg:
                if dlg.ShowModal() == wx.ID_OK:
                    return
        num_rows = len(df)
        with DeduplicateIndexDialog(self, [c for c in self.colnames
                                           if c not in self.drugs_col]) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                data = df
                if dlg.isSortDate.GetValue():
                    data = data.sort_values(config.Read('DateCol'), ascending=True)
                if dlg.keys:
                    data = data.drop_duplicates(subset=[self.colnames[k] for k in dlg.keys],
                                                keep='first')
                if num_rows == len(data):
                    message = 'No duplicates found.'
                else:
                    message = f'{ num_rows - len(data) } duplicates were removed.'
                with wx.MessageDialog(self, message, 'Deduplication Finished', style=wx.OK) as dlg:
                    dlg.ShowModal()
            else:
                return
        keys = []
        for c in self.colnames:
            if c not in self.drugs_col:
                keys.append(c)
        columns = [c for c in self.colnames if c not in self.drugs_col] + ['GENUS', 'SPECIES', 'GRAM']
        if self.identifier_col in columns:
            columns.remove(self.identifier_col)
        if self.date_col in columns:
            columns.remove(self.date_col)

        with BiogramIndexDialog(self, columns,
                                start=data[self.date_col].min(),
                                end=data[self.date_col].max()) as dlg:
            if dlg.ShowModal() == wx.ID_OK and dlg.indexes:
                # filter data within the date range
                data = data[(data[self.date_col].dt.date >= dlg.startDate.GetValue())
                            & (data[self.date_col].dt.date <= dlg.endDate.GetValue())]
                BiogramGeneratorThread(data,
                                       self.date_col,
                                       self.identifier_col,
                                       self.organism_col,
                                       dlg.indexes,
                                       keys,
                                       dlg.includeCount.GetValue(),
                                       dlg.includePercent.GetValue(),
                                       dlg.includeNarstStyle.GetValue(),
                                       columns,
                                       self.drug_data)
                progress_bar = PulseProgressBarDialog('Generating Antibiogram', 'Calculating...')
            else:
                return

    def write_output(self, sens, resists, biogram_sens, biogram_resists, biogram_narst_s,
                     identifier_col):
        with wx.FileDialog(self, "Please select the output file for your antibiogram",
                           wildcard="Excel file (*xlsx)|*xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            file_path = file_dialog.GetPath()
            if os.path.splitext(file_path)[1] != '.xlsx':
                file_path = file_path + '.xlsx'
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                if sens is not None:
                    sens[identifier_col].to_excel(writer, sheet_name='count_S')
                if resists is not None:
                    resists[identifier_col].to_excel(writer, sheet_name='count_R')
                if biogram_sens is not None:
                    biogram_sens[identifier_col].to_excel(writer, sheet_name='percent_S')
                if biogram_resists is not None:
                    biogram_resists[identifier_col].to_excel(writer, sheet_name='percent_R')
                if biogram_narst_s is not None:
                    biogram_narst_s[identifier_col].to_excel(writer, sheet_name='narst_s')
        except:
            with wx.MessageDialog(self, 'Failed', 'Antibiogram Generator', style=wx.OK) as dlg:
                dlg.ShowModal()
        else:
            with wx.MessageDialog(self, 'Output Saved.', 'Antibiogram Generator', style=wx.OK) as dlg:
                dlg.ShowModal()



class GenApp(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__(self, redirect, filename)

    def OnInit(self):
        # create frame here
        global config
        config = wx.Config('Mivisor')
        frame = MainFrame()
        self.SetTopWindow(frame)
        frame.Show()
        return True
