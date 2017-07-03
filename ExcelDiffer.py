#coding=utf-8
"""
运行主程序
"""
from GUI.MainWindow import *
import sys
app = QtGui.QApplication(sys.argv)
mainWindow = QtGui.QMainWindow()
ui = MainWindow()
ui.setupUi(mainWindow)
mainWindow.show()
sys.exit(app.exec_())