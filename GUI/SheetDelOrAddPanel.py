#coding=utf-8
from PyQt4 import QtCore, QtGui
import string

red = QtGui.QColor(255, 0, 0)
deepBlue = QtGui.QColor(0, 0, 139)

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

class SheetDelOrAddPanel(QtGui.QWidget):
    def __init__(self,parent = None):
        super(SheetDelOrAddPanel,self).__init__(parent)
        self.setupUi()
    def setupUi(self):
        self.setObjectName(_fromUtf8("SheetDelOrAddPanel"))
        self.resize(553, 363)
        self.verticalLayout = QtGui.QVBoxLayout(self)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))

        self.widget = QtGui.QWidget()
        self.widget.setObjectName(_fromUtf8("widget"))
        self.verticalLayout.addWidget(self.widget)

        self.horizontalLayout = QtGui.QHBoxLayout(self.widget)
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.splitter = QtGui.QSplitter(self.widget)
        self.splitter.setOrientation(QtCore.Qt.Vertical)
        self.splitter.setObjectName(_fromUtf8("splitter"))
        self.label = QtGui.QLabel(self.splitter)
        self.label.setObjectName(_fromUtf8("label"))
        self.tableWidget = QtGui.QTableWidget(self.splitter)
        self.tableWidget.setObjectName(_fromUtf8("tableWidget"))
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.tableWidget.setStyleSheet("QTableWidget { border: 0px solid black }; ")

        item = QtGui.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtGui.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        self.horizontalLayout.addWidget(self.splitter)
        self.tabWidget_2 = QtGui.QTabWidget(self.widget)
        self.tabWidget_2.setObjectName(_fromUtf8("tabWidget_2"))

        self.horizontalLayout.addWidget(self.tabWidget_2)
        self.label.raise_()
        self.tableWidget.raise_()
        self.tabWidget_2.raise_()
        self.tableWidget.raise_()
        self.splitter.raise_()

        self.retranslateUi()

        self.tabWidget_2.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(self)

        self.tableWidget.setSelectionMode(QtGui.QAbstractItemView.NoSelection)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()    #调整table大小

    def retranslateUi(self):
        self.setWindowTitle(_translate("self", "self", None))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("self", "改动", None))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("self", "sheet名", None))


    def setSheetDelOrAddInfo(self,fileName1,fileName2,addSheets,delSheets):

        fontHead = QtGui.QFont()
        fontHead.setBold(True)
        fontHead.setWeight(75)


        def setMatrixInTable(fileName,tableName,matrix):
            nCol = 0
            nRow = len(matrix)
            if(nRow > 0):
                nCol = len(matrix[0])

            frame = QtGui.QFrame()
            layout = QtGui.QVBoxLayout()
            frame.setLayout(layout)
            tableWidget = QtGui.QTableWidget()
            tableWidget.setObjectName(_fromUtf8("tableWidget"))

            label = QtGui.QLabel()
            label.setText(_translate("SheetDelOrAddPanel", fileName, None))

            layout.addWidget(label)
            layout.addWidget(tableWidget)

            tableWidget.setColumnCount(nCol)
            tableWidget.setRowCount(nRow)

            tableWidget.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)

            for col in range(nCol):
                item = QtGui.QTableWidgetItem()
                tableWidget.setHorizontalHeaderItem(col, item)
                head = string.uppercase[col]
                item.setText(_translate("DisplayPanel", head, None))
            for row in range(nRow):
                item = QtGui.QTableWidgetItem()
                tableWidget.setVerticalHeaderItem(row, item)
                item.setText(_translate("DisplayPanel", str(row + 1), None))

            for i in range(nRow):
                for j in range(nCol):
                    item = QtGui.QTableWidgetItem()
                    tableWidget.setItem(i, j, item)
                    value = self.codeUTF8(matrix[i][j])
                    item.setText(_translate("DisplayPanel", value, None))  #存入数据

            self.tabWidget_2.addTab(frame, _fromUtf8(tableName))


        strMessageLabel = "共计删除{0}个sheet，增加{1}个sheet".format(len(addSheets),len(delSheets))
        self.label.setText(_translate("SheetDelOrAddPanel", strMessageLabel, None))
        self.tableWidget.setRowCount(len(addSheets) + len(delSheets))

        for index,value in enumerate(addSheets):
            sheetName,matrix = value
            item = QtGui.QTableWidgetItem()
            self.tableWidget.setItem(index, 0, item)
            item.setText(_translate("SheetDelOrAddPanel", "新增", None))
            item.setTextColor(deepBlue)

            addLink = "<a href=\"{0}\">{1}</a>".format(index, self.codeUTF8(sheetName))
            hyperlinkLabelAddSheet = QtGui.QLabel(_translate("SheetDelOrAddPanel", addLink, None), self)
            self.connect(hyperlinkLabelAddSheet, QtCore.SIGNAL("linkActivated(QString)"),
                         self.labelSheetClicked)

            hyperlinkLabelAddSheet.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter | QtCore.Qt.AlignCenter)
            self.tableWidget.setCellWidget(index, 1, hyperlinkLabelAddSheet)

            setMatrixInTable(fileName2,sheetName,matrix)




        for index,value in enumerate(delSheets):
            sheetName,matrix = value
            item = QtGui.QTableWidgetItem()
            self.tableWidget.setItem(index + len(addSheets), 0, item)
            item.setText(_translate("SheetDelOrAddPanel", "删除", None))
            item.setTextColor(red)

            addLink = "<a href=\"{0}\">{1}</a>".format(index + len(addSheets), self.codeUTF8(sheetName))
            hyperlinkLabelDelSheet = QtGui.QLabel(_translate("SheetDelOrAddPanel", addLink, None), self)

            self.connect(hyperlinkLabelDelSheet, QtCore.SIGNAL("linkActivated(QString)"),
                         self.labelSheetClicked)

            hyperlinkLabelDelSheet.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter | QtCore.Qt.AlignCenter)
            self.tableWidget.setCellWidget(index + len(addSheets), 1, hyperlinkLabelDelSheet)

            setMatrixInTable(fileName1,sheetName, matrix)

    def labelSheetClicked(self,QString):
        index, boolean = QString.toInt()
        self.tabWidget_2.setCurrentIndex(index)



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


