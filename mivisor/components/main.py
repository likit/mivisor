import wx
import os
import pandas as pd
from ObjectListView import ObjectListView, ColumnDefn, FastObjectListView
from wx.lib.mixins.listctrl import CheckListCtrlMixin


class DataRow(object):
    def __init__(self, id, series) -> None:
        self.id = id
        for k, v in series.items():
            setattr(self, k, v)

    def to_list(self, columns):
        return [getattr(self, c) for c in columns]


class BiogramIndexDialog(wx.Dialog):
    def __init__(self, parent, columns, title='Biogram Indexes'):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.indexes = []
        self.choices = columns
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        self.chlbox = wx.CheckListBox(self, choices=columns)
        self.chlbox.Bind(wx.EVT_CHECKLISTBOX, self.on_checked)
        self.index_items_list = wx.ListCtrl(self, wx.ID_ANY, style=wx.LC_LIST)
        main_sizer.Add(self.chlbox, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.index_items_list, 1, wx.ALL | wx.EXPAND, 10)
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
            self.index_items_list.Append([self.choices[item]])


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
        super(DrugListCtrl, self).__init__(parent, style=wx.LC_REPORT)
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


class MainPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)

        self.df = None
        self.colnames = []
        self.organism_col = config.Read('OrganismCol', '')
        self.identifier_col = config.Read('IdentifierCol', '')
        self.date_col = config.Read('DateCol', '')
        self.specimens_col = config.Read('SpecimensCol', '')
        self.drugs_col = config.Read('Drugs', '').split(';') or []
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        load_button = wx.Button(self, label="Load")
        save_button = wx.Button(self, label="Save")
        add_button = wx.Button(self, label="Add Column")
        copy_button = wx.Button(self, label="Copy Column")
        config_btn = wx.Button(self, label='Config')
        melt_btn = wx.Button(self, label='Melt')
        load_btn = wx.Button(self, label='Load Drugs')
        generate_btn = wx.Button(self, label='Generate')
        load_button.Bind(wx.EVT_BUTTON, self.load_data_from_file)
        save_button.Bind(wx.EVT_BUTTON, self.save_records)
        add_button.Bind(wx.EVT_BUTTON, self.add_column)
        copy_button.Bind(wx.EVT_BUTTON, self.copy_column)
        config_btn.Bind(wx.EVT_BUTTON, self.configure)
        melt_btn.Bind(wx.EVT_BUTTON, self.melt)
        generate_btn.Bind(wx.EVT_BUTTON, self.generate)
        load_btn.Bind(wx.EVT_BUTTON, self.load_drug_registry)

        self.dataOlv = FastObjectListView(self, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        self.dataOlv.oddRowsBackColor = wx.Colour(230, 230, 230, 100)
        self.dataOlv.evenRowsBackColor = wx.WHITE
        self.dataOlv.cellEditMode = ObjectListView.CELLEDIT_DOUBLECLICK
        main_sizer.Add(self.dataOlv, 1, wx.ALL | wx.EXPAND, 10)
        btn_sizer.Add(load_button, 0, wx.ALL, 5)
        btn_sizer.Add(save_button, 0, wx.ALL, 5)
        btn_sizer.Add(add_button, 0, wx.ALL, 5)
        btn_sizer.Add(copy_button, 0, wx.ALL, 5)
        btn_sizer.Add(config_btn, 0, wx.ALL, 5)
        btn_sizer.Add(melt_btn, 0, wx.ALL, 5)
        btn_sizer.Add(load_btn, 0, wx.ALL, 5)
        btn_sizer.Add(generate_btn, 0, wx.ALL, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALL, 5)
        self.SetSizer(main_sizer)
        self.Fit()


    def load_data_from_file(self, event):
        self.df = pd.read_excel('tiny-sample.xlsx')
        self.df = self.df.dropna(how='all').fillna('')
        self.setColumns()
        self.data = [DataRow(idx, row) for idx, row in self.df.iterrows()]
        self.dataOlv.SetObjects(self.data)


    def setColumns(self):
        columns = []
        for c in self.df.columns:
            self.colnames.append(c)
            col_type = str(self.df.dtypes.get(c))
            if col_type.startswith('int') or col_type.startswith('float'):
                formatter = '%.1f'
            elif col_type.startswith('datetime'):
                formatter = '%s'
            else:
                formatter = '%s'
            columns.append(
                ColumnDefn(title=c.title(), align='left', stringConverter=formatter, valueGetter=c, minimumWidth=50))
        self.dataOlv.SetColumns(columns)

    def save_records(self, event):
        data = [row.to_list(self.colnames) for row in self.data]
        ids = [row.id for row in self.data]
        df = pd.DataFrame(data=data, index=ids, columns=self.colnames)
        dask_df = dd.from_pandas(df, npartitions=2)
        print(dask_df)

    def add_column(self, event):
        self.dataOlv.AddColumnDefn(ColumnDefn(
            title='strain',
            align='left',
            valueGetter='strain',
            minimumWidth=50,
        ))
        for row in self.data:
            row.strain = 'alpha'
        self.dataOlv.RepopulateList()

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
                self.identifier_col = self.colnames[dlg.identifier_combo_ctrl.GetSelection()]
                self.date_col = self.colnames[dlg.date_combo_ctrl.GetSelection()]
                self.organism_col = self.colnames[dlg.organism_combo_ctrl.GetSelection()]
                self.specimens_col = self.colnames[dlg.specimens_combo_ctrl.GetSelection()]
                self.drugs_col = dlg.drug_listctrl.drugs
                config.Write('IdentifierCol', self.identifier_col)
                config.Write('DateCol', self.date_col)
                config.Write('OrganismCol', self.organism_col)
                config.Write('SpecimensCol', self.specimens_col)
                config.Write('Drugs', ';'.join(self.drugs_col))

    def melt(self, event):
        self.keys = []
        for c in self.colnames:
            if c not in self.drugs_col:
                self.keys.append(c)
        data = [row.to_list(self.colnames) for row in self.data]
        ids = [row.id for row in self.data]
        df = pd.DataFrame(data=data, index=ids, columns=self.colnames)
        self.melted_df = df.melt(id_vars=self.keys)
        print(self.melted_df.columns)
        print('done melting..')

    def generate(self, event):
        columns = [c for c in self.colnames if c not in self.drugs_col] + ['GENUS', 'SPECIES', 'GRAM']
        if self.identifier_col in columns:
            columns.remove(self.identifier_col)
        if self.date_col in columns:
            columns.remove(self.date_col)
        with BiogramIndexDialog(self, columns) as dlg:
            if dlg.ShowModal() == wx.ID_OK and dlg.indexes:
                organism_df = pd.read_excel('organisms2020.xlsx')
                _melted_df = pd.merge(self.melted_df, organism_df, how='inner')
                _melted_df = pd.merge(_melted_df, self.drug_data, right_on='abbr', left_on='variable', how='outer')
                indexes = [columns[idx] for idx in dlg.indexes]
                total = _melted_df.pivot_table(index=indexes, columns=['group', 'variable'], aggfunc='count')
                sens = _melted_df[_melted_df['value'] == 'S'].pivot_table(index=indexes,
                                                                                  columns=['group', 'variable'],
                                                                                  aggfunc='count')
                biogram = (sens/total*100).applymap(lambda x: round(x, 2))
                formatted_total = total.applymap(lambda x: '' if pd.isna(x) else '{:.0f}'.format(x))
                biogram_narst_s = biogram.fillna('-').applymap(str) + " (" + formatted_total + ")"
                biogram_narst_s = biogram_narst_s.applymap(lambda x: '' if x.startswith('-') else x)
                biogram_narst_s[self.identifier_col].to_excel('biogram.xlsx')

    def load_drug_registry(self, event):
        try:
            drug_df = pd.read_json(os.path.join('drugs.json'))
        except:
            return pd.DataFrame(columns=['drug', 'abbreviation', 'group'])
        else:
            if drug_df.empty:
                drug_df = pd.DataFrame(columns=['drug', 'abbreviation', 'group'])
            else:
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


class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY,
                          title="ObjectListView Demo", size=(800, 600))
        panel = MainPanel(self)
        self.statusbar = self.CreateStatusBar(2)
        self.statusbar.SetStatusText('The app is ready to roll.')
        self.statusbar.SetStatusText('The is for the analytics information', 1)


class GenApp(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__(self, redirect, filename)

    def OnInit(self):
        # create frame here
        global config
        config = wx.Config('Mivisor')
        frame = MainFrame()
        frame.Show()
        return True

