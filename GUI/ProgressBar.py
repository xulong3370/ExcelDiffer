#coding=utf-8
"""
进度条
"""
from PyQt4 import QtCore, QtGui
import sys

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)


class ProgressBar(QtGui.QWidget):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)
        screen = QtGui.QDesktopWidget().screenGeometry()
        self.setFixedSize(400,100)
        self.setObjectName(_fromUtf8("ProgressBar"))
        self.setWindowTitle(_translate("ProgressBar", "ProgressBar", None))

        self.cancelButton = QtGui.QPushButton(u"终止")

        self.progressBar = QtGui.QProgressBar(self)

        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(self.cancelButton)

        self.label = QtGui.QLabel(u'准备中...')

        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(self.label,0,QtCore.Qt.AlignCenter)
        vbox.addWidget(self.progressBar)
        vbox.addLayout(hbox)
        self.setLayout(vbox)
        self.progressBar.setValue(0)

    def updateValue(self,i):
        """
        更新数据
        :param i:
        :return:
        """
        self.progressBar.setValue(i)
    def updateWindowTitle(self,QString):
        """
        跟新sheet进度
        :param QString:
        :return:
        """
        self.setWindowTitle(QString)

    def updateMessage(self,QString):
        """
        跟新算法进度
        :param QString:
        :return:
        """
        self.label.setText(QString)
    def setRange(self,min,max):
        """
        设定范围
        :param min:
        :param max:
        :return:
        """
        self.progressBar.setRange(min,max)
