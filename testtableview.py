from PyQt5 import QtCore, QtGui, QtWidgets


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        table_view = QtWidgets.QTableView()
        self.setCentralWidget(table_view)

        model = QtGui.QStandardItemModel(5, 5, self)
        table_view.setModel(model)

        for i in range(model.rowCount()):
            for j in range(model.columnCount()):
                it = QtGui.QStandardItem(f"{i}-{j}")
                model.setItem(i, j, it)

        selection_model = table_view.selectionModel()
        selection_model.selectionChanged.connect(self.on_selectionChanged)

    # @QtCore.pyqtSlot('QItemSelection', 'QItemSelection')
    def on_selectionChanged(self, selected, deselected):
        print("selected: ")
        for ix in selected.indexes():
            print(ix.data())

        print("deselected: ")
        for ix in deselected.indexes():
            print(ix.data())


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())