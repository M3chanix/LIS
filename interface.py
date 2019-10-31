import sys
from PyQt5 import QtWidgets, QtCore, QtGui
import class_desc
import design_v5
import pandas as pd

Qt = QtCore.Qt


class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return QtCore.QVariant(str(
                    self._data.values[index.row()][index.column()]))
        return QtCore.QVariant()


class ExampleApp(QtWidgets.QMainWindow, design_v5.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.proc = class_desc.Processor("C:/py1/LIS")
        self.connect_view_model()

    def connect_view_model(self):
        dep_dict = self.proc.department_dict.copy()
        tabWidget = ExampleApp.findChildren(self, QtWidgets.QTabWidget)[0]
        for department_name, index in zip(dep_dict, range(1, tabWidget.count())):
            tabWidget.setTabText(index, department_name)
            toolbox = tabWidget.widget(index).findChildren(QtWidgets.QToolBox)[0]
            tableview1 = toolbox.widget(0).findChildren(QtWidgets.QTableView)[0]
            tableview2 = toolbox.widget(1).findChildren(QtWidgets.QTableView)[0]
            tableview3 = toolbox.widget(2).findChildren(QtWidgets.QTableView)[0]
            model1 = PandasModel(self.proc.__dict__[dep_dict[department_name]].processed_3_5)
            model2 = PandasModel(self.proc.__dict__[dep_dict[department_name]].processed_6_7)
            model3 = PandasModel(self.proc.__dict__[dep_dict[department_name]].processed_2_4)
            tableview1.setModel(model1)
            tableview2.setModel(model2)
            tableview3.setModel(model3)
            header = tableview3.horizontalHeader()
            header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)

        # Инициализация перовой странички с таблицами с общими для всего института данными
        tabWidget.setTabText(0, "Общее")
        toolbox = tabWidget.widget(0).findChildren(QtWidgets.QToolBox)[0]
        tableview1 = toolbox.widget(0).findChildren(QtWidgets.QTableView)[0]
        tableview2 = toolbox.widget(1).findChildren(QtWidgets.QTableView)[0]
        model1 = PandasModel(self.proc.processed_1_2)
        model2 = PandasModel(self.proc.processed_2_3)
        tableview1.setModel(model1)
        tableview2.setModel(model2)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ExampleApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
