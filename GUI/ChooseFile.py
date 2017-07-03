#coding=utf-8
"""
载入file 界面
"""

from IO.ExcelRW import *
import xlrd

from PyQt4 import QtCore, QtGui

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

class ChooseFile(QtGui.QWidget):
    def __init__(self,parent = None):
        super(ChooseFile, self).__init__(parent)
        self.setupUi()
    def setupUi(self):
        self.workbook1 = None
        self.workbook2 = None


        self.setObjectName(_fromUtf8("ChooseFile"))
        #self.resize(609, 393)
        self.gridLayout = QtGui.QGridLayout(self)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.frameExcel = QtGui.QFrame(self)
        self.frameExcel.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frameExcel.setFrameShadow(QtGui.QFrame.Raised)
        self.frameExcel.setObjectName(_fromUtf8("frameExcel"))
        self.gridLayout_2 = QtGui.QGridLayout(self.frameExcel)
        self.gridLayout_2.setObjectName(_fromUtf8("gridLayout_2"))
        self.labelExcel = QtGui.QLabel(self.frameExcel)
        self.labelExcel.setObjectName(_fromUtf8("labelExcel"))
        self.gridLayout_2.addWidget(self.labelExcel, 0, 0, 1, 1)
        self.labelSheetName = QtGui.QLabel(self.frameExcel)
        self.labelSheetName.setObjectName(_fromUtf8("labelSheetName"))
        self.gridLayout_2.addWidget(self.labelSheetName, 2, 0, 1, 1)
        self.lineEditFile = QtGui.QLineEdit(self.frameExcel)
        self.lineEditFile.setObjectName(_fromUtf8("lineEditFile"))
        self.gridLayout_2.addWidget(self.lineEditFile, 1, 1, 1, 1)
        self.LoadFile = QtGui.QPushButton(self.frameExcel)
        self.LoadFile.setObjectName(_fromUtf8("LoadFile"))
        self.gridLayout_2.addWidget(self.LoadFile, 1, 2, 1, 1)
        self.labelFileName = QtGui.QLabel(self.frameExcel)
        self.labelFileName.setObjectName(_fromUtf8("labelFileName"))
        self.gridLayout_2.addWidget(self.labelFileName, 1, 0, 1, 1)
        self.comboBoxSheetName = QtGui.QComboBox(self.frameExcel)
        self.comboBoxSheetName.setObjectName(_fromUtf8("comboBoxSheetName"))
        self.gridLayout_2.addWidget(self.comboBoxSheetName, 2, 1, 1, 1)
        self.gridLayout.addWidget(self.frameExcel, 0, 0, 1, 1)
        self.frameExcel_2 = QtGui.QFrame(self)
        self.frameExcel_2.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frameExcel_2.setFrameShadow(QtGui.QFrame.Raised)
        self.frameExcel_2.setObjectName(_fromUtf8("frameExcel_2"))
        self.gridLayout_3 = QtGui.QGridLayout(self.frameExcel_2)
        self.gridLayout_3.setObjectName(_fromUtf8("gridLayout_3"))
        self.labelExcel_2 = QtGui.QLabel(self.frameExcel_2)
        self.labelExcel_2.setObjectName(_fromUtf8("labelExcel_2"))
        self.gridLayout_3.addWidget(self.labelExcel_2, 0, 0, 1, 1)
        self.lineEditFile_2 = QtGui.QLineEdit(self.frameExcel_2)
        self.lineEditFile_2.setObjectName(_fromUtf8("lineEditFile_2"))
        self.gridLayout_3.addWidget(self.lineEditFile_2, 1, 1, 1, 1)
        self.LoadFile_2 = QtGui.QPushButton(self.frameExcel_2)
        self.LoadFile_2.setObjectName(_fromUtf8("LoadFile_2"))
        self.gridLayout_3.addWidget(self.LoadFile_2, 1, 2, 1, 1)
        self.labelSheetName_2 = QtGui.QLabel(self.frameExcel_2)
        self.labelSheetName_2.setObjectName(_fromUtf8("labelSheetName_2"))
        self.gridLayout_3.addWidget(self.labelSheetName_2, 2, 0, 1, 1)
        self.comboBoxSheetName_2 = QtGui.QComboBox(self.frameExcel_2)
        self.comboBoxSheetName_2.setObjectName(_fromUtf8("comboBoxSheetName_2"))
        self.gridLayout_3.addWidget(self.comboBoxSheetName_2, 2, 1, 1, 1)
        self.labelFileName_2 = QtGui.QLabel(self.frameExcel_2)
        self.labelFileName_2.setObjectName(_fromUtf8("labelFileName_2"))
        self.gridLayout_3.addWidget(self.labelFileName_2, 1, 0, 1, 1)
        self.gridLayout.addWidget(self.frameExcel_2, 1, 0, 1, 1)
        self.verticalLayout = QtGui.QVBoxLayout()
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))
        self.pushButtonCompare = QtGui.QPushButton(self)
        self.pushButtonCompare.setObjectName(_fromUtf8("pushButtonCompare"))
        self.verticalLayout.addWidget(self.pushButtonCompare)
        self.pushButtonAutoCompare = QtGui.QPushButton(self)
        self.pushButtonAutoCompare.setObjectName(_fromUtf8("pushButtonAutoCompare"))
        self.verticalLayout.addWidget(self.pushButtonAutoCompare)
        self.gridLayout.addLayout(self.verticalLayout, 1, 1, 1, 1)

        self.retranslateUi()
        #QtCore.QMetaObject.connectSlotsByName()

        self.adjustSize()
        self.setMinimumWidth(609)

        self.LoadFile.clicked.connect(self.loadFile1Dialog)  #链接所有函数槽
        self.LoadFile_2.clicked.connect(self.loadFile2Dialog)
        self.pushButtonCompare.clicked.connect(self.compare)
        self.pushButtonAutoCompare.clicked.connect(self.compareAll)

    def retranslateUi(self):
        self.setWindowTitle(_translate("ChooseFile", "ChooseFile", None))
        self.labelExcel.setText(_translate("ChooseFile", "Excel1", None))
        self.labelSheetName.setText(_translate("ChooseFile", "sheet名：", None))
        self.LoadFile.setText(_translate("ChooseFile", "选择文件", None))
        self.labelFileName.setText(_translate("ChooseFile", "文件名：", None))
        self.labelExcel_2.setText(_translate("ChooseFile", "Excel2", None))
        self.LoadFile_2.setText(_translate("ChooseFile", "选择文件", None))
        self.labelSheetName_2.setText(_translate("ChooseFile", "sheet名：", None))
        self.labelFileName_2.setText(_translate("ChooseFile", "文件名：", None))
        self.pushButtonCompare.setText(_translate("ChooseFile", "单个对比", None))
        self.pushButtonAutoCompare.setText(_translate("ChooseFile", "自动全部对比", None))


    def loadFile1Dialog(self):
        """
        excel1文件选择
        :return: workbook1
        """
        try:
            fileName= QtGui.QFileDialog.getOpenFileName(self,  "选取文件",  "C:/","All Files (*);;Excel Files (*.xlsx)")
            self.lineEditFile.clear()
            fileNameUtf8Byte = fileName.toUtf8()
            fileNameStr = str(fileNameUtf8Byte)
            self.comboBoxSheetName.clear()

            if(not fileNameStr or
                       str(self.lineEditFile_2.text().toUtf8()) != fileNameStr):
                self.workbook1 = ExtractExcel(fileNameStr)
                self.lineEditFile.setText(fileName)
                for sheets in self.workbook1.getExcelSheetsNameList():
                    self.comboBoxSheetName.addItem(sheets)
            else:
                reply = QtGui.QMessageBox.question(self, u'警告', u"文件路径相同，依然要选择吗？",
                                                        QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                if(reply == QtGui.QMessageBox.Yes):
                    self.workbook1 = ExtractExcel(fileNameStr)
                    self.lineEditFile.setText(fileName)
                    for sheets in self.workbook1.getExcelSheetsNameList():
                        self.comboBoxSheetName.addItem(sheets)
        except IOError:
            pass   #未选择
        except xlrd.XLRDError:
            self.chooseErrorDialog()

    def loadFile2Dialog(self):
        """
        excel1文件选择
        :return: workbook1
        """
        try:
            fileName= QtGui.QFileDialog.getOpenFileName(self,  "选取文件",  "C:/","All Files (*);;Excel Files (*.xlsx)")
            self.lineEditFile_2.clear()
            fileNameUtf8Byte = fileName.toUtf8()
            fileNameStr = str(fileNameUtf8Byte)
            self.comboBoxSheetName_2.clear()

            if (not fileNameStr or
                        str(self.lineEditFile.text().toUtf8()) != fileNameStr):
                self.workbook2 = ExtractExcel(fileNameStr)
                self.lineEditFile_2.setText(fileName)
                for sheets in self.workbook2.getExcelSheetsNameList():
                    self.comboBoxSheetName_2.addItem(sheets)
            else:
                reply = QtGui.QMessageBox.question(self, u'警告', u"文件路径相同，依然要选择吗？",
                                                        QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                if (reply == QtGui.QMessageBox.Yes):
                    self.workbook2 = ExtractExcel(fileNameStr)
                    self.lineEditFile_2.setText(fileName)
                    for sheets in self.workbook2.getExcelSheetsNameList():
                        self.comboBoxSheetName_2.addItem(sheets)
        except IOError:
            pass   #未选择
        except xlrd.XLRDError:
            self.chooseErrorDialog()


    def chooseErrorDialog(self):
        reply = QtGui.QMessageBox.warning(self, u'错误',
                                          u"文件类型错误或文件已损坏", u"确定")

    def compare(self):
        """
        对比选中的
        :return:
        """
        if(self.workbook1 != None and self.workbook2 != None and
                   self.comboBoxSheetName.currentIndex() != -1 and self.comboBoxSheetName_2.currentIndex() != -1):
            choose1 = str(self.comboBoxSheetName.currentText().toUtf8())
            choose2 = str(self.comboBoxSheetName_2.currentText().toUtf8())
            self.sheets1 = choose1
            self.sheets2 = choose2
            self.emit(QtCore.SIGNAL('executeThread()'))
        else:
            reply = QtGui.QMessageBox.warning(self, u'错误',
                                              u"未选择sheet", u"确定")

    def compareAll(self):
        """
        对比所有sheet
        :return:
        """
        if(self.workbook1 != None and self.workbook2 != None):
            sheets1 = self.workbook1.getExcelSheetsNameList()
            sheets2 = self.workbook2.getExcelSheetsNameList()
            if(sheets1 != None and sheets2 != None):
                self.sheets1 = sheets1
                self.sheets2 = sheets2
                self.emit(QtCore.SIGNAL('executeThread()'))
        else:
            reply = QtGui.QMessageBox.warning(self, u'错误',
                                              u"无sheet", u"确定")

    def codeUTF8(self,s):
        """
        中文兼容转码
        :param s:
        :return:
        """
        if isinstance(s, unicode):
            value = s.encode('utf8')
        else:
            value = str(s)
        return value

