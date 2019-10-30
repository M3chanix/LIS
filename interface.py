import sys
from PyQt5 import QtWidgets, QtCore, QtGui
Qt = QtCore.Qt
import design
import class_desc


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


class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.proc = class_desc.Processor("C:/py1/LIS")
        self.proc.process_1_2()
        self.proc.process_2_3()
        self.proc.process_2_4()
        self.model = PandasModel(self.proc.processed_2_3)
        self.tableView.setModel(self.model)



def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ExampleApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
