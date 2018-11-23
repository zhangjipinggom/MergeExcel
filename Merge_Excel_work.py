#ZJP
#Merge_Excel_work.py  19:13

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtWidgets, QtGui
import pandas as pd
from MergeExcel2 import Ui_MainWindow

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.setWindowIcon(QtGui.QIcon("icon_mainwindow.png"))
        QtWidgets.QToolTip.setFont(QFont('Arial',10))
        self.pushButton_table1.clicked.connect(self.on_actionOpen_triggered)
        self.pushButton_table2.clicked.connect(self.on_actionOpen_triggered2)
        self.pushButton_merge.clicked.connect(self.merge)
        self.radioButton_2.toggled.connect(self.set_pattern2)
        self.radioButton.toggled.connect(self.set_pattern1)
        self.textEdit_key.setText("学号")



    @pyqtSlot()
    def on_actionOpen_triggered(self):
        import functools
        my_file_path = QFileDialog.getOpenFileName(self, "open file", './')
        self.table1 = my_file_path[0]
        path_c = self.table1.split("/")
        self.path_save = functools.reduce(lambda x, y: x+'/'+y, path_c[0:-1])
        self.textBrowser_table1.setText(self.table1)

    def on_actionOpen_triggered2(self):
        my_file_path = QFileDialog.getOpenFileName(self, "open file", './')
        self.table2 = my_file_path[0]
        self.textBrowser_table2.setText(self.table2)

    def merge(self):
        self.textBrowser_results.setText("begin")
        data1 = pd.read_excel(self.table1)
        data2 = pd.read_excel(self.table2)
        self.key = self.textEdit_key.toPlainText()
        self.textBrowser_results.setText("2")
        data3 = pd.merge(data1, data2, how="left", on=[self.key])
        self.textBrowser_results.setText("3")
        saved_file = "{}/merge.xls".format(self.path_save)
        data3.to_excel(saved_file)
        self.textBrowser_results.setText("merged file is saved to {}. \n{} is merged into {} according to {}".format
                                         (saved_file, self.table2, self.table1, self.key))

    def set_pattern1(self):
        a = 1

    def set_pattern2(self):
        a = 2


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.processEvents()  # 使程序在显示启动画面的同时仍能响应鼠标其他事件
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec_())
