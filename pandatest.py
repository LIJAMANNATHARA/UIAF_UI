import sys
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QTableWidget, QApplication,QTableView
Qt = QtCore.Qt
import pandas as pd

# class PandasModel(QtCore.QAbstractTableModel):
#     def __init__(self, data,headerdata, parent=None):
#         QtCore.QAbstractTableModel.__init__(self, parent)
#         self._data = data
#         self.headerdata=headerdata
#
#     def rowCount(self, parent=None):
#         return len(self._data.values)
#
#     def columnCount(self, parent=None):
#         return self._data.columns.size
#
#     def data(self, index, role=Qt.DisplayRole):
#         if index.isValid():
#             if role == Qt.DisplayRole:
#                 return QtCore.QVariant(str(
#                     self._data.values[index.row()][index.column()]))
#         return QtCore.QVariant()
#
#     # def setData(self, index, value, role=Qt.EditRole):
#     #     if index.isValid():
#     #         row = index.row()
#     #         col = index.column()
#     #         self._data.loc[row][col] = str(value)
#     #         if self.data(index,QtCore.Qt.DisplayRole)==value:
#     #             self.dataChanged.emit(index, index, (Qt.DisplayRole,))
#     #         return True
#     #     return False
#     def setData(self, index, value, role):
#         if not index.isValid():
#             return False
#         if role != QtCore.Qt.EditRole:
#             return False
#         row = index.row()
#         if row < 0 or row >= len(self._data.values):
#             return False
#         column = index.column()
#         if column < 0 or column >= self._data.columns.size:
#             return False
#         self._data.values[row][column] = str(value)
#         self.dataChanged.emit(index, index,(Qt.DisplayRole,))
#         return True
#
#     def flags(self, index):
#         return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
#
#     def headerData(self, col, orientation, role):
#         if orientation == Qt.Horizontal and role == Qt.DisplayRole:
#             return QtCore.QVariant(self.headerdata[col])
#         return QtCore.QVariant()
class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data,headerData, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data
        self.headerdata = headerData

    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if index.isValid():
            if role == QtCore.Qt.DisplayRole or role == QtCore.Qt.EditRole:
                return QtCore.QVariant(
                    self._data.iloc[index.row(),index.column()])
        return QtCore.QVariant()

    def headerData(self, col, orientation, role):

     if orientation == Qt.Horizontal and role == Qt.DisplayRole:
                 return QtCore.QVariant(self.headerdata[col])
     return QtCore.QVariant()

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEditable

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        if index.isValid():
            self._data.iloc[index.row(),index.column()] = value
            if self.data(index,QtCore.Qt.DisplayRole) == value:
                self.dataChanged.emit(index, index)
                return True
        return QtCore.QVariant()




