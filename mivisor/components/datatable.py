import wx
import pandas
from wx.grid import GridTableBase


class DataTable(GridTableBase):
    def __init__(self):
        super(DataTable, self).__init__()
        self.df = pandas.DataFrame()
        self.odd = wx.grid.GridCellAttr()
        self.odd.SetBackgroundColour('blue')
        self.odd.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD))

        self.even = wx.grid.GridCellAttr()
        self.even.SetBackgroundColour('white')
        self.even.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD))

    def GetNumberRows(self):
        return len(self.df)

    def GetNumberCols(self):
        return len(self.df.columns.values)

    def IsEmptyCell(self, row, col):
        return pandas.isnull(self.df.iloc[row, col])

    def GetValue(self, row, col):
        value = self.df.iloc[row, col]
        if not pandas.isnull(value):
            return value
        else:
            return None

    def SetValue(self, row, col, value):
        self.df.iloc[row, col] = value


class DataGrid(wx.grid.Grid):
    def __init__(self, parent):
        super(DataGrid, self).__init__(parent)
        self.table = DataTable()

    def set_table(self, df):
        self.table.df = df
        self.SetTable(self.table, True)
