import os
import sys
import json
import sqlite3
from datetime import datetime

import wx
import wx.adv
import pandas as pd
import xlsxwriter

if hasattr(wx, 'ItemAttr'):
    wx.ListItemAttr = wx.ItemAttr

from ObjectListView import ObjectListView, ColumnDefn, FastObjectListView
from threading import Thread
from pubsub import pub

from components.drug_dialog import DrugRegFormDialog


CLOSE_PROGRESS_BAR_SIGNAL = 'close-progressbar'
WRITE_TO_EXCEL_FILE_SIGNAL = 'write-to-excel-file'
ENABLE_BUTTONS = 'enable-buttons'
DISABLE_BUTTONS = 'disable-buttons'
DATABASE_SCHEMA_VERSION = 1


def patch_object_list_view():
    if getattr(ObjectListView, '_mivisor_size_patch', False):
        return

    def _handle_size(self, evt):
        self._PossibleFinishCellEdit()
        evt.Skip()
        self._ResizeSpaceFillingColumns()
        if not hasattr(self, 'stEmptyListMsg') or not self.stEmptyListMsg:
            return

        sz = self.GetClientSize()
        x = 0
        y = int(sz.GetHeight() / 3)
        width = int(sz.GetWidth())
        height = int(sz.GetHeight())
        if 'phoenix' in wx.PlatformInfo:
            self.stEmptyListMsg.SetSize(x, y, width, height)
        else:
            self.stEmptyListMsg.SetDimensions(x, y, width, height)

    ObjectListView._HandleSize = _handle_size
    ObjectListView._mivisor_size_patch = True


patch_object_list_view()


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

    @staticmethod
    def _organism_lookup():
        organism_df = pd.read_excel(os.path.join('appdata', 'organisms2020.xlsx'))
        return organism_df[['ORGANISM', 'GENUS', 'SPECIES', 'GRAM']]

    def _empty_result(self, indexes):
        empty_index = pd.MultiIndex.from_arrays([[] for _ in indexes], names=indexes)
        empty_columns = pd.MultiIndex.from_arrays(
            [[], [], []], names=[None, 'group', 'variable']
        )
        return pd.DataFrame(index=empty_index, columns=empty_columns)

    def _wrap_result(self, frame):
        return pd.concat({self.identifier_col: frame}, axis=1)

    @staticmethod
    def _coerce_numeric(frame):
        return frame.apply(pd.to_numeric, errors='coerce')

    @staticmethod
    def _format_count(value):
        if pd.isna(value):
            return ''
        try:
            return '{:.0f}'.format(float(value))
        except (TypeError, ValueError):
            return ''

    def _build_outputs(self, long_df, indexes):
        if long_df.empty:
            total = self._empty_result(indexes)
            sens = self._empty_result(indexes)
            resists = self._empty_result(indexes)
        else:
            working_df = long_df.copy()
            working_df['is_s'] = (working_df['value'] == 'S').astype('int64')
            working_df['is_resist'] = working_df['value'].isin(['I', 'R']).astype('int64')

            grouped = working_df.groupby(indexes + ['group', 'variable'], observed=True)[
                [self.identifier_col, 'is_s', 'is_resist']
            ].agg({
                self.identifier_col: 'count',
                'is_s': 'sum',
                'is_resist': 'sum',
            })

            total = self._wrap_result(grouped[self.identifier_col].unstack(['group', 'variable']))
            sens = self._wrap_result(grouped['is_s'].unstack(['group', 'variable']))
            resists = self._wrap_result(grouped['is_resist'].unstack(['group', 'variable']))

        total = self._coerce_numeric(total)
        sens = self._coerce_numeric(sens)
        resists = self._coerce_numeric(resists)
        biogram_resists = (resists / total * 100).round(2)
        biogram_sens = (sens / total * 100).round(2)
        formatted_total = total.map(self._format_count)
        biogram_narst_s = biogram_sens.fillna('-').map(str) + " (" + formatted_total + ")"
        biogram_narst_s = biogram_narst_s.map(lambda x: '' if x.startswith('-') else x)
        return sens, resists, biogram_sens, biogram_resists, biogram_narst_s

    def run(self):
        indexes = [self.columns[idx] for idx in self.indexes]
        drug_columns = [column for column in self.data.columns if column not in self.keys]

        long_df = pd.DataFrame(columns=indexes + ['group', 'variable', 'value', self.identifier_col])
        if drug_columns:
            organism_df = self._organism_lookup().rename(columns={'ORGANISM': self.organism_col})
            annotated_df = self.data.merge(organism_df, on=self.organism_col, how='inner')
            melted_df = annotated_df.melt(id_vars=self.keys + ['GENUS', 'SPECIES', 'GRAM'],
                                          value_vars=drug_columns)
            drug_lookup = self.drug_data[['abbr', 'group']].drop_duplicates()
            merged_df = melted_df.merge(drug_lookup, left_on='variable', right_on='abbr', how='inner')
            long_df = merged_df[[
                *indexes,
                self.identifier_col,
                'group',
                'variable',
                'value',
            ]]

        sens, resists, biogram_sens, biogram_resists, biogram_narst_s = self._build_outputs(long_df, indexes)
        wx.CallAfter(pub.sendMessage, CLOSE_PROGRESS_BAR_SIGNAL)
        wx.CallAfter(pub.sendMessage,
                     WRITE_TO_EXCEL_FILE_SIGNAL,
                     sens=sens if self.include_count else None,
                     resists=resists if self.include_count else None,
                     biogram_sens=biogram_sens if self.include_percent else None,
                     biogram_resists=biogram_resists if self.include_percent else None,
                     biogram_narst_s=biogram_narst_s if self.include_narst else None,
                     identifier_col=self.identifier_col)


class DatabaseBiogramGeneratorThread(BiogramGeneratorThread):
    def __init__(self, facts_df, identifier_col, indexes, include_count, include_percent, include_narst):
        self.facts_df = facts_df
        super().__init__(
            data=pd.DataFrame(),
            date_col='',
            identifier_col=identifier_col,
            organism_col='',
            indexes=list(range(len(indexes))),
            keys=[],
            include_count=include_count,
            include_percent=include_percent,
            include_narst=include_narst,
            columns=indexes,
            drug_data=pd.DataFrame(),
        )

    def run(self):
        indexes = [self.columns[idx] for idx in self.indexes]
        long_df = self.facts_df.rename(columns={
            'drug_group': 'group',
            'drug': 'variable',
            'sensitivity': 'value',
        })
        sens, resists, biogram_sens, biogram_resists, biogram_narst_s = self._build_outputs(long_df, indexes)
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


def format_datetime(value):
    if pd.isna(value):
        return ''
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value)


def to_wx_date(value):
    if pd.isna(value) or value is None:
        return wx.DateTime.Now()
    if hasattr(value, 'to_pydatetime'):
        value = value.to_pydatetime()
    if hasattr(value, 'year') and hasattr(value, 'month') and hasattr(value, 'day'):
        return wx.DateTime.FromDMY(value.day, value.month - 1, value.year)
    return wx.DateTime.Now()


class HeatmapConfigDialog(wx.Dialog):
    def __init__(self, parent, fields, start=None, end=None, title='Heatmap Configuration'):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        form_sizer = wx.FlexGridSizer(3, 2, 10, 10)

        form_sizer.Add(wx.StaticText(self, label='Row Field'))
        self.field_choice = wx.Choice(self, choices=fields)
        if fields:
            self.field_choice.SetSelection(0)
        form_sizer.Add(self.field_choice, 0, wx.EXPAND)

        form_sizer.Add(wx.StaticText(self, label='Start'))
        self.startDate = wx.adv.DatePickerCtrl(self, dt=start or wx.DateTime.Now())
        form_sizer.Add(self.startDate, 0, wx.EXPAND)

        form_sizer.Add(wx.StaticText(self, label='End'))
        self.endDate = wx.adv.DatePickerCtrl(self, dt=end or wx.DateTime.Now())
        form_sizer.Add(self.endDate, 0, wx.EXPAND)

        cutoff_sizer = wx.BoxSizer(wx.HORIZONTAL)
        cutoff_sizer.Add(wx.StaticText(self, label='Minimum Isolates'), 0, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.ncutoff = wx.SpinCtrl(self, min=0, initial=0)
        cutoff_sizer.Add(self.ncutoff, 0)

        btn_sizer = self.CreateStdDialogButtonSizer(wx.OK | wx.CANCEL)
        main_sizer.Add(form_sizer, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(cutoff_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)
        self.SetSizer(main_sizer)
        self.Fit()


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

        configured_drugs = {drug for drug in config.Read('Drugs').split(';') if drug}
        detected_drugs = self.detect_drug_columns()

        for idx, col in enumerate(self.cols):
            if col in configured_drugs or col in detected_drugs:
                self.CheckItem(idx)

    def detect_drug_columns(self):
        try:
            drug_df = pd.read_json(os.path.join('appdata', 'drugs.json'))
        except:
            return set()

        abbreviations = set()
        for abbreviation in drug_df.get('abbreviation', pd.Series(dtype='object')).dropna():
            if isinstance(abbreviation, str):
                abbreviations.update(
                    a.strip().upper() for a in abbreviation.split(',') if a.strip()
                )

        detected = set()
        for col in self.cols:
            if isinstance(col, str) and col.strip().upper() in abbreviations:
                detected.add(col)
        return detected

    def on_check(self, event):
        item = event.GetItem()
        idx = item.GetId()
        col = self.cols[idx]
        if self.IsItemChecked(idx):
            if col not in self.drugs:
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
        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY, title="Mivisor Version 2021.1", size=(800, 600))
        panel = wx.Panel(self)
        # TODO: figure out how to update the statusbar's text from the frame's children
        self.statusbar = self.CreateStatusBar(2)
        self.statusbar.SetStatusText('The app is ready to roll.')
        self.statusbar.SetStatusText('This is for the analytics information', 1)
        menuBar = wx.MenuBar()
        fileMenu = wx.Menu()
        registryMenu = wx.Menu()
        databaseMenu = wx.Menu()
        menuBar.Append(fileMenu, '&File')
        menuBar.Append(registryMenu, 'Re&gistry')
        menuBar.Append(databaseMenu, '&Database')
        loadItem = fileMenu.Append(wx.ID_ANY, 'Load Data', 'Load Data')
        exportItem = fileMenu.Append(wx.ID_ANY, 'Export Data', 'Export Data')
        fileMenu.AppendSeparator()
        fileItem = fileMenu.Append(wx.ID_EXIT, '&Quit', 'Quit Application')
        drugItem = registryMenu.Append(wx.ID_ANY, 'Drugs', 'Drug Registry')
        exportDatabaseItem = databaseMenu.Append(wx.ID_ANY, 'Save Database', 'Save current data to a database')
        generateDatabaseItem = databaseMenu.Append(wx.ID_ANY, 'Generate Antibiogram', 'Generate antibiogram from a database')
        heatmapDatabaseItem = databaseMenu.Append(wx.ID_ANY, 'Generate Heatmap', 'Generate heatmap from a database')
        self.SetMenuBar(menuBar)
        self.Bind(wx.EVT_MENU, lambda x: self.Close(), fileItem)
        self.Bind(wx.EVT_MENU, self.open_drug_dialog, drugItem)
        self.Bind(wx.EVT_MENU, self.export_data, exportItem)
        self.Bind(wx.EVT_MENU, self.open_load_data_dialog, loadItem)
        self.Bind(wx.EVT_MENU, self.export_database, exportDatabaseItem)
        self.Bind(wx.EVT_MENU, self.generate_from_database, generateDatabaseItem)
        self.Bind(wx.EVT_MENU, self.generate_heatmap_from_database, heatmapDatabaseItem)

        self.Bind(wx.EVT_CLOSE, self.OnClose)
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
        self.current_data_path = ''
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
        if df.empty:
            with wx.MessageDialog(self, 'No data to export. Please load data first.',
                                  'Export Data', style=wx.OK) as dlg:
                dlg.ShowModal()
                return
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
            drug_df = pd.DataFrame()
        if drug_df.empty:
            drug_df = pd.DataFrame(columns=['drug', 'abbreviation', 'group'])

        drug_list = []
        drug_df = drug_df.sort_values(['group'])
        for idx, row in drug_df.iterrows():
            abbreviation = row.get('abbreviation', '')
            if isinstance(abbreviation, str) and abbreviation.strip():
                abbrs = [a.strip().upper() for a in abbreviation.split(',') if a.strip()]
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
            self.current_data_path = filepath
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

    def require_configuration(self):
        if not all([self.date_col, self.identifier_col, self.organism_col]):
            self.configure(None)
        if not all([self.date_col, self.identifier_col, self.organism_col]):
            with wx.MessageDialog(self, 'Please configure identifier, date, and organism columns first.',
                                  'Configuration Required', style=wx.OK) as dlg:
                dlg.ShowModal()
            return False
        if not self.drugs_col:
            with wx.MessageDialog(self, 'Please configure at least one drug column first.',
                                  'Configuration Required', style=wx.OK) as dlg:
                dlg.ShowModal()
            return False
        return True

    def build_current_dataframe(self):
        return pd.DataFrame([d.to_dict(self.colnames) for d in self.data])

    @staticmethod
    def normalize_sensitivity(value):
        if pd.isna(value):
            return ''
        return str(value).strip().upper()

    def load_organism_lookup(self):
        organism_df = pd.read_excel(os.path.join('appdata', 'organisms2020.xlsx'))
        return organism_df[['ORGANISM', 'GENUS', 'SPECIES', 'GRAM']]

    def build_database_profile(self):
        return {
            'schema_version': DATABASE_SCHEMA_VERSION,
            'identifier_col': self.identifier_col,
            'date_col': self.date_col,
            'organism_col': self.organism_col,
            'specimens_col': self.specimens_col,
            'drugs_col': self.drugs_col,
            'colnames': self.colnames,
        }

    def build_database_facts(self, df):
        drug_columns = [col for col in self.drugs_col if col in df.columns]
        if not drug_columns:
            return pd.DataFrame()

        facts_df = df.copy()
        facts_df['record_id'] = range(len(facts_df))
        organism_lookup = self.load_organism_lookup().rename(columns={'ORGANISM': self.organism_col})
        facts_df = facts_df.merge(organism_lookup, on=self.organism_col, how='left')
        for col in ['GENUS', 'SPECIES', 'GRAM']:
            if col in facts_df:
                facts_df[col] = facts_df[col].fillna('')
        facts_df['organism_name'] = (
            facts_df['GENUS'].astype(str).str.strip() + ' ' + facts_df['SPECIES'].astype(str).str.strip()
        ).str.strip()
        facts_df.loc[facts_df['organism_name'] == '', 'organism_name'] = facts_df[self.organism_col].astype(str)

        id_vars = [col for col in facts_df.columns if col not in drug_columns]
        melted_df = facts_df.melt(id_vars=id_vars, value_vars=drug_columns,
                                  var_name='drug', value_name='sensitivity')
        melted_df['sensitivity'] = melted_df['sensitivity'].map(self.normalize_sensitivity)
        melted_df = melted_df[melted_df['sensitivity'] != '']

        drug_lookup = self.drug_data[['abbr', 'group']].drop_duplicates().rename(
            columns={'abbr': 'drug', 'group': 'drug_group'}
        )
        melted_df = melted_df.merge(drug_lookup, on='drug', how='inner')
        melted_df['added_at'] = datetime.utcnow().isoformat(timespec='seconds')
        return melted_df

    @staticmethod
    def build_database_metadata(profile_json, source_path):
        return pd.DataFrame([{
            'schema_version': DATABASE_SCHEMA_VERSION,
            'created_at': datetime.utcnow().isoformat(timespec='seconds'),
            'source_path': source_path or '',
            'profile_json': profile_json,
        }])

    def export_database(self, event):
        if not self.require_configuration():
            return

        df = self.build_current_dataframe()
        if df.empty:
            with wx.MessageDialog(self, 'No data to export. Please load data first.',
                                  'Save Database', style=wx.OK) as dlg:
                dlg.ShowModal()
                return

        facts_df = self.build_database_facts(df)
        if facts_df.empty:
            with wx.MessageDialog(self, 'No database rows could be created from the configured drug columns.',
                                  'Save Database', style=wx.OK) as dlg:
                dlg.ShowModal()
                return

        with wx.FileDialog(self, "Please select the database file",
                           wildcard="SQLite file (*.sqlite;*.db)|*.sqlite;*.db",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            file_path = file_dialog.GetPath()
            if os.path.splitext(file_path)[1] not in ('.sqlite', '.db'):
                file_path = file_path + '.sqlite'

        metadata_df = self.build_database_metadata(
            json.dumps(self.build_database_profile()),
            self.current_data_path,
        )
        try:
            with sqlite3.connect(file_path) as con:
                facts_df.to_sql('facts', con=con, if_exists='replace', index=False)
                metadata_df.to_sql('metadata', con=con, if_exists='replace', index=False)
        except:
            with wx.MessageDialog(self, 'Failed to save database.',
                                  'Save Database', style=wx.OK) as dlg:
                dlg.ShowModal()
        else:
            with wx.MessageDialog(self, 'Database saved.',
                                  'Save Database', style=wx.OK) as dlg:
                dlg.ShowModal()

    def read_database(self, file_path):
        with sqlite3.connect(file_path) as con:
            facts_df = pd.read_sql_query('SELECT * FROM facts', con)
            metadata_df = pd.read_sql_query('SELECT * FROM metadata', con)
        if metadata_df.empty:
            raise ValueError('metadata is empty')
        profile = json.loads(metadata_df.iloc[-1]['profile_json'])
        return facts_df, metadata_df, profile

    def prepare_database_facts(self, facts_df, profile):
        working_df = facts_df.copy()
        date_col = profile.get('date_col', '')
        if date_col and date_col in working_df.columns:
            working_df[date_col] = pd.to_datetime(working_df[date_col], errors='coerce')
        if 'organism_name' not in working_df.columns:
            if 'GENUS' in working_df.columns and 'SPECIES' in working_df.columns:
                working_df['organism_name'] = (
                    working_df['GENUS'].fillna('').astype(str).str.strip()
                    + ' ' +
                    working_df['SPECIES'].fillna('').astype(str).str.strip()
                ).str.strip()
                organism_col = profile.get('organism_col', '')
                if organism_col and organism_col in working_df.columns:
                    working_df.loc[working_df['organism_name'] == '', 'organism_name'] = (
                        working_df[organism_col].astype(str)
                    )
            else:
                organism_col = profile.get('organism_col', '')
                if organism_col and organism_col in working_df.columns:
                    working_df['organism_name'] = working_df[organism_col].astype(str)
        return working_df

    def deduplicate_database_facts(self, facts_df, profile):
        non_drug_columns = [
            col for col in facts_df.columns
            if col not in {'record_id', 'drug', 'drug_group', 'sensitivity', 'added_at'}
        ]
        with DeduplicateIndexDialog(self, non_drug_columns) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return None
            filtered_facts = facts_df
            date_col = profile.get('date_col', '')
            if dlg.isSortDate.GetValue() and date_col in filtered_facts.columns:
                filtered_facts = filtered_facts.sort_values(date_col, ascending=True)
            records_df = filtered_facts[non_drug_columns + ['record_id']].drop_duplicates('record_id')
            if dlg.keys:
                selected_keys = [non_drug_columns[k] for k in dlg.keys]
                deduped_records = records_df.drop_duplicates(subset=selected_keys, keep='first')
            else:
                deduped_records = records_df
            removed = len(records_df) - len(deduped_records)
            with wx.MessageDialog(self,
                                  'No duplicates found.' if removed == 0 else f'{removed} duplicates were removed.',
                                  'Deduplication Finished', style=wx.OK) as msg_dlg:
                msg_dlg.ShowModal()
            return filtered_facts[filtered_facts['record_id'].isin(deduped_records['record_id'])]

    def generate_from_database(self, event):
        with wx.FileDialog(self, "Select a database",
                           wildcard="SQLite file (*.sqlite;*.db)|*.sqlite;*.db",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            file_path = file_dialog.GetPath()

        try:
            facts_df, metadata_df, profile = self.read_database(file_path)
        except:
            with wx.MessageDialog(self, 'Failed to read database.',
                                  'Database', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        facts_df = self.prepare_database_facts(facts_df, profile)
        facts_df = self.deduplicate_database_facts(facts_df, profile)
        if facts_df is None:
            return

        date_col = profile.get('date_col', '')
        identifier_col = profile.get('identifier_col', '')
        if not identifier_col or identifier_col not in facts_df.columns:
            with wx.MessageDialog(self, 'Database metadata is missing the identifier column.',
                                  'Database', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        columns = [
            col for col in facts_df.columns
            if col not in {'record_id', 'drug', 'drug_group', 'sensitivity', 'added_at'}
        ]
        if identifier_col in columns:
            columns.remove(identifier_col)
        if date_col in columns:
            columns.remove(date_col)

        start = to_wx_date(facts_df[date_col].min()) if date_col in facts_df.columns else wx.DateTime.Now()
        end = to_wx_date(facts_df[date_col].max()) if date_col in facts_df.columns else wx.DateTime.Now()
        with BiogramIndexDialog(self, columns, start=start, end=end) as dlg:
            if dlg.ShowModal() != wx.ID_OK or not dlg.indexes:
                return
            filtered_facts = facts_df
            if date_col in filtered_facts.columns:
                start_date = pd.Timestamp(dlg.startDate.GetValue().FormatISODate()).date()
                end_date = pd.Timestamp(dlg.endDate.GetValue().FormatISODate()).date()
                filtered_facts = filtered_facts[
                    (filtered_facts[date_col].dt.date >= start_date)
                    & (filtered_facts[date_col].dt.date <= end_date)
                ]
            DatabaseBiogramGeneratorThread(
                filtered_facts[[*columns, identifier_col, 'drug_group', 'drug', 'sensitivity']],
                identifier_col,
                [columns[idx] for idx in dlg.indexes],
                dlg.includeCount.GetValue(),
                dlg.includePercent.GetValue(),
                dlg.includeNarstStyle.GetValue(),
            )
            PulseProgressBarDialog('Generating Antibiogram', f'Calculating from {os.path.basename(file_path)}...')

    def create_heatmap_dataframe(self, facts_df, row_field, organism_name, identifier_col, cutoff=0):
        filtered_df = facts_df[facts_df['organism_name'] == organism_name].copy()
        if filtered_df.empty:
            return pd.DataFrame()

        filtered_df['is_s'] = (filtered_df['sensitivity'] == 'S').astype('int64')
        grouped = filtered_df.groupby([row_field, 'drug'], observed=True)[
            [identifier_col, 'is_s']
        ].agg({
            identifier_col: 'count',
            'is_s': 'sum',
        })
        counts = grouped[identifier_col].unstack('drug')
        sens = grouped['is_s'].unstack('drug')
        if cutoff > 0:
            counts = counts.where(counts >= cutoff)
        return ((sens / counts) * 100).round(2)

    def plot_heatmap(self, df, title):
        import numpy as np
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import seaborn as sns

        plot_df = df.replace(r'^\s*$', np.nan, regex=True)
        plot_df = plot_df.dropna(axis=1, how='all').dropna(axis=0, how='all')
        if plot_df.empty:
            with wx.MessageDialog(self, 'The plot could not be created because the data table is empty.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        default_dir = os.path.dirname(self.current_data_path) if self.current_data_path else os.getcwd()
        file_path = os.path.join(default_dir, 'heatmap.png')
        with wx.TextEntryDialog(self, 'Enter the output PNG path',
                                'Save Heatmap', value=file_path) as path_dialog:
            if path_dialog.ShowModal() != wx.ID_OK:
                return
            file_path = path_dialog.GetValue().strip()
            if not file_path:
                return

        if os.path.splitext(file_path)[1] != '.png':
            file_path = file_path + '.png'
        output_dir = os.path.dirname(file_path) or '.'
        if not os.path.isdir(output_dir):
            with wx.MessageDialog(self, 'The selected output folder does not exist.',
                                  'Heatmap', style=wx.OK) as dlg:
                return
                dlg.ShowModal()
            return

        try:
            cluster = sns.clustermap(plot_df.fillna(120),
                                     cmap=sns.diverging_palette(20, 220, n=7),
                                     linewidths=0.2)
            cluster.fig.suptitle(title)
            cluster.savefig(file_path)
            plt.close(cluster.fig)
        except:
            with wx.MessageDialog(self, 'The plot could not be generated or saved.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()
        else:
            with wx.MessageDialog(self, 'Heatmap saved.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()

    def generate_heatmap_from_database(self, event):
        with wx.FileDialog(self, "Select a database",
                           wildcard="SQLite file (*.sqlite;*.db)|*.sqlite;*.db",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return
            file_path = file_dialog.GetPath()

        try:
            facts_df, metadata_df, profile = self.read_database(file_path)
        except:
            with wx.MessageDialog(self, 'Failed to read database.',
                                  'Database', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        facts_df = self.prepare_database_facts(facts_df, profile)
        identifier_col = profile.get('identifier_col', '')
        date_col = profile.get('date_col', '')
        if not identifier_col or identifier_col not in facts_df.columns:
            with wx.MessageDialog(self, 'Database metadata is missing the identifier column.',
                                  'Database', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        heatmap_fields = [
            col for col in facts_df.columns
            if col not in {'record_id', 'drug', 'drug_group', 'sensitivity', 'added_at',
                           'organism_name', identifier_col, date_col}
        ]
        if not heatmap_fields:
            with wx.MessageDialog(self, 'No fields are available to build heatmap rows.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        start = to_wx_date(facts_df[date_col].min()) if date_col in facts_df.columns else wx.DateTime.Now()
        end = to_wx_date(facts_df[date_col].max()) if date_col in facts_df.columns else wx.DateTime.Now()
        with HeatmapConfigDialog(self, heatmap_fields, start=start, end=end) as dlg:
            if dlg.ShowModal() != wx.ID_OK or dlg.field_choice.GetSelection() == wx.NOT_FOUND:
                return
            row_field = heatmap_fields[dlg.field_choice.GetSelection()]
            cutoff = dlg.ncutoff.GetValue()
            filtered_facts = facts_df
            if date_col in filtered_facts.columns:
                start_date = pd.Timestamp(dlg.startDate.GetValue().FormatISODate()).date()
                end_date = pd.Timestamp(dlg.endDate.GetValue().FormatISODate()).date()
                filtered_facts = filtered_facts[
                    (filtered_facts[date_col].dt.date >= start_date)
                    & (filtered_facts[date_col].dt.date <= end_date)
                ]

        organisms = sorted([name for name in filtered_facts['organism_name'].dropna().unique() if str(name).strip()])
        if not organisms:
            with wx.MessageDialog(self, 'No organisms are available for heatmap generation.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()
            return

        with wx.SingleChoiceDialog(self, "Select an organism", "Heatmap Organism", organisms) as org_dlg:
            if org_dlg.ShowModal() != wx.ID_OK:
                return
            organism_name = org_dlg.GetStringSelection()

        heatmap_df = self.create_heatmap_dataframe(filtered_facts, row_field, organism_name, identifier_col, cutoff)
        if heatmap_df.empty:
            with wx.MessageDialog(self, 'No heatmap data could be generated for the selected organism and field.',
                                  'Heatmap', style=wx.OK) as dlg:
                dlg.ShowModal()
            return
        self.plot_heatmap(heatmap_df, f'{organism_name} by {row_field}')

    def setColumns(self):
        columns = []
        self.colnames = []
        for c in self.df.columns:
            self.colnames.append(c)
            col_type = str(self.df.dtypes.get(c))
            if col_type.startswith('int') or col_type.startswith('float'):
                formatter = '%.1f'
            elif col_type.startswith('datetime'):
                formatter = format_datetime
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
            self.configure(None)
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
                                start=to_wx_date(data[self.date_col].min()),
                                end=to_wx_date(data[self.date_col].max())) as dlg:
            if dlg.ShowModal() == wx.ID_OK and dlg.indexes:
                # filter data within the date range
                start_date = pd.Timestamp(dlg.startDate.GetValue().FormatISODate()).date()
                end_date = pd.Timestamp(dlg.endDate.GetValue().FormatISODate()).date()
                data = data[(data[self.date_col].dt.date >= start_date)
                            & (data[self.date_col].dt.date <= end_date)]
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
