#coding=utf-8
"""
展示界面，包含一个tab 两个 excel展示表格
"""
from PyQt4 import QtCore, QtGui
from Algorithm.DifferAlgorithm import *
from Algorithm.AnalysisMatrixInfo import AnalysisMatrixInfo
from IO.ExcelRW import *
import sys

red = QtGui.QColor(255, 0, 0)
blue = QtGui.QColor(135, 206, 255)
yellow = QtGui.QColor(255, 255, 0)
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

class DisplayPanel(QtGui.QWidget):
    def __init__(self,parent = None):
        super(DisplayPanel,self).__init__(parent)
        self.setupUi()

    def setupUi(self):
        screen = QtGui.QDesktopWidget().screenGeometry() #获得screen
        self.setObjectName(_fromUtf8("DisplayPanel"))
        self.resize(screen.width(), screen.height())

        self.verticalLayout = QtGui.QVBoxLayout(self)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))

        self.label = QtGui.QLabel()
        self.label2 = QtGui.QLabel()

        self.horizontalLayout2 = QtGui.QHBoxLayout()
        self.horizontalLayout2.setObjectName(_fromUtf8("horizontalLayout2"))
        self.horizontalLayout2.addWidget(self.label)
        self.horizontalLayout2.addWidget(self.label2)

        self.verticalLayout.addLayout(self.horizontalLayout2)

        self.initTopExcelChangeDisplayUI()  #形成顶部excelUI
        self.verticalLayout.addLayout(self.horizontalLayout)


        self.initBottomTabChangeDisplayUI() #底部tabUI
        self.verticalLayout.addWidget(self.tabWidget)

        self.retranslateUi()
        self.tabWidget.setCurrentIndex(0)
        #QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self):
        self.setWindowTitle(_translate("DisplayPanel", "DisplayPanel", None))

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.rowChange), _translate("DisplayPanel", "行增删", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.colChange), _translate("DisplayPanel", "列增删", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.cellChange), _translate("DisplayPanel", "单元格改动", None))

    def initTopExcelChangeDisplayUI(self):
        """
        顶部excelUI
        :return:
        """
        self.horizontalLayout = QtGui.QHBoxLayout()
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.preMatrix = QtGui.QTableWidget()        #修改前矩阵table
        self.preMatrix.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.preMatrix.setObjectName(_fromUtf8("preMatrix"))
        self.preMatrix.setStyleSheet("selection-background-color:blue;")
        self.horizontalLayout.addWidget(self.preMatrix)

        self.postMatrix = QtGui.QTableWidget()       #修改后矩阵table
        self.postMatrix.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.postMatrix.setObjectName(_fromUtf8("postMatrix"))
        self.postMatrix.setStyleSheet("selection-background-color:gray;")
        self.horizontalLayout.addWidget(self.postMatrix)

        #链接函数槽
        self.connect(self.preMatrix, QtCore.SIGNAL("itemClicked (QTableWidgetItem*)"),
                                   self.matrixTableClicked)
        self.connect(self.postMatrix, QtCore.SIGNAL("itemClicked (QTableWidgetItem*)"),
                                   self.matrixTableClicked)
        self.connect(self.preMatrix.horizontalHeader(), QtCore.SIGNAL("sectionClicked(int)"),
                                  self.matrixTableColSelected)
        self.connect(self.preMatrix.verticalHeader(), QtCore.SIGNAL("sectionClicked(int)"),
                                   self.matrixTableRowSelected)
        self.connect(self.postMatrix.horizontalHeader(), QtCore.SIGNAL("sectionClicked(int)"),
                                  self.matrixTableColSelected)
        self.connect(self.postMatrix.verticalHeader(), QtCore.SIGNAL("sectionClicked(int)"),
                                   self.matrixTableRowSelected)


    def initBottomTabChangeDisplayUI(self):
        """
        底部tapUI
        :return:
        """
        self.tabWidget = QtGui.QTabWidget()       #bottom 的 tab
        self.tabWidget.setObjectName(_fromUtf8("tabWidget"))
        self.colChange = QtGui.QWidget()
        self.colChange.setObjectName(_fromUtf8("colChange"))
        self.tabWidget.addTab(self.colChange, _fromUtf8(""))
        self.rowChange = QtGui.QWidget()
        self.rowChange.setObjectName(_fromUtf8("rowChange"))
        self.tabWidget.addTab(self.rowChange, _fromUtf8(""))
        self.cellChange = QtGui.QWidget()
        self.cellChange.setObjectName(_fromUtf8("cellChange"))
        self.tabWidget.addTab(self.cellChange, _fromUtf8(""))

        self.verticalLayoutColChange = QtGui.QVBoxLayout(self.colChange)  #tab中设定格式
        self.verticalLayoutColChange.setObjectName(_fromUtf8("verticalLayoutColChange"))
        self.verticalLayoutRowChange = QtGui.QVBoxLayout(self.rowChange)
        self.verticalLayoutRowChange.setObjectName(_fromUtf8("verticalLayoutRowChange"))
        self.verticalLayoutCellChange = QtGui.QVBoxLayout(self.cellChange)
        self.verticalLayoutCellChange.setObjectName(_fromUtf8("verticalLayoutCellChange"))

        self.initTabMessageUI()  #创建tab中的message
        self.initTabTableUI()  #创建tab中的table

    def initTabMessageUI(self):
        """
        三个在不同tab中的 text
        :return:
        """
        self.labelColChange = QtGui.QLabel(self.colChange)
        self.verticalLayoutColChange.addWidget(self.labelColChange)
        self.labelColChange.setTextFormat(QtCore.Qt.AutoText)
        self.labelColChange.setObjectName(_fromUtf8("labelColChange"))

        self.labelRowChange = QtGui.QLabel(self.rowChange)
        self.verticalLayoutRowChange.addWidget(self.labelRowChange)
        self.labelRowChange.setTextFormat(QtCore.Qt.AutoText)
        self.labelRowChange.setObjectName(_fromUtf8("labelRowChange"))

        self.labelCellChange = QtGui.QLabel(self.cellChange)
        self.verticalLayoutCellChange.addWidget(self.labelCellChange)
        self.labelCellChange.setTextFormat(QtCore.Qt.AutoText)
        self.labelCellChange.setObjectName(_fromUtf8("labelCellChange"))

        self.labelColChange.setText(_translate("DisplayPanel", "共计新增0列，删除0列", None))
        self.labelRowChange.setText(_translate("DisplayPanel", "共计新增0行，删除0行", None))
        self.labelCellChange.setText(_translate("DisplayPanel", "共计0个单元格改动", None))

    def initTabTableUI(self):
        fontHead = QtGui.QFont()
        fontHead.setBold(True)
        fontHead.setWeight(75)

        self.tabColChangeTable = QtGui.QTableWidget(self.colChange)
        self.tabColChangeTable.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)

        self.verticalLayoutColChange.addWidget(self.tabColChangeTable)
        self.tabColChangeTable.setStyleSheet("QTableWidget {border: 0px solid black }; ")
        self.tabColChangeTable.setObjectName(_fromUtf8("tabColChangeTable"))
        self.tabColChangeTable.setColumnCount(2)
        item = QtGui.QTableWidgetItem()
        self.tabColChangeTable.setHorizontalHeaderItem(0, item)
        item.setText(_translate("DisplayPanel", "改动", None))
        item.setFont(fontHead)
        item = QtGui.QTableWidgetItem()
        item.setText(_translate("DisplayPanel", "列号", None))
        item.setFont(fontHead)
        self.tabColChangeTable.setHorizontalHeaderItem(1, item)
        self.tabColChangeTable.verticalHeader().setVisible(False)

        self.tabRowChangeTable = QtGui.QTableWidget(self.rowChange)
        self.tabRowChangeTable.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.verticalLayoutRowChange.addWidget(self.tabRowChangeTable)
        self.tabRowChangeTable.setStyleSheet("QTableWidget { border: 0px solid black }; ")
        self.tabRowChangeTable.setObjectName(_fromUtf8("tabRowChangeTable"))
        self.tabRowChangeTable.setColumnCount(2)
        item = QtGui.QTableWidgetItem()
        self.tabRowChangeTable.setHorizontalHeaderItem(0, item)
        item.setText(_translate("DisplayPanel", "改动", None))
        item.setFont(fontHead)
        item = QtGui.QTableWidgetItem()
        item.setText(_translate("DisplayPanel", "行号", None))
        item.setFont(fontHead)
        self.tabRowChangeTable.setHorizontalHeaderItem(1, item)
        self.tabRowChangeTable.verticalHeader().setVisible(False)

        self.tabCellChangeTable = QtGui.QTableWidget(self.cellChange)
        self.tabCellChangeTable.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.verticalLayoutCellChange.addWidget(self.tabCellChangeTable)
        self.tabCellChangeTable.setStyleSheet("QTableWidget { border: 0px solid black }; ")

        self.tabCellChangeTable.setObjectName(_fromUtf8("tabCellChangeTable"))
        self.tabCellChangeTable.setColumnCount(3)
        item = QtGui.QTableWidgetItem()
        self.tabCellChangeTable.setHorizontalHeaderItem(0, item)
        item.setText(_translate("DisplayPanel", "坐标", None))
        item.setFont(fontHead)
        item = QtGui.QTableWidgetItem()
        self.tabCellChangeTable.setHorizontalHeaderItem(1, item)
        item.setFont(fontHead)
        item.setText(_translate("DisplayPanel", "旧值", None))
        item = QtGui.QTableWidgetItem()
        self.tabCellChangeTable.setHorizontalHeaderItem(2, item)
        item.setFont(fontHead)
        item.setText(_translate("DisplayPanel", "新值", None))
        self.tabCellChangeTable.verticalHeader().setVisible(False)

        self.tabColChangeTable.setSelectionMode(QtGui.QAbstractItemView.NoSelection )  #tab的不能选中，不然会很难看
        self.tabCellChangeTable.setSelectionMode(QtGui.QAbstractItemView.NoSelection)
        self.tabRowChangeTable.setSelectionMode(QtGui.QAbstractItemView.NoSelection)

    def setResultInGUI(self,matrixInfo):
        """
        将数据展示
        :param matrixInfo: 数据信息矩阵
        :return:
        """
        preMatrix = matrixInfo.preMatrix
        postMatrix = matrixInfo.postMatrix

        delRowInfo = matrixInfo.delRowInfo
        delColInfo = matrixInfo.delColInfo
        addRowInfo = matrixInfo.addRowInfo
        addColInfo = matrixInfo.addColInfo
        modifyCellInfo = matrixInfo.modifyCellInfo

        nRow = matrixInfo.maxNRow
        nCol = matrixInfo.maxNCol
        
        preRowHead = matrixInfo.preRowHeadList
        preColHead = matrixInfo.preColHeadList
        postRowHead = matrixInfo.postRowHeadList
        postColHead = matrixInfo.postColHeadList
        
        def setInGuiTable(tableWidget,matrix,rowHead,colHead):
            """
            输入表格控件 载入样式和数据
            :param tableWidget: 表格控件
            :param matrix:
            :param rowHead:
            :param colHead:
            :return:
            """
            tableWidget.setColumnCount(nCol)
            tableWidget.setRowCount(nRow)

            for index, value in enumerate(colHead):
                item = QtGui.QTableWidgetItem()
                tableWidget.setVerticalHeaderItem(index, item)
                value = value if value != 'CHANGED'else ''
                item.setText(_translate("DisplayPanel", str(value), None))


            for index, value in enumerate(rowHead):
                item = QtGui.QTableWidgetItem()
                tableWidget.setHorizontalHeaderItem(index, item)
                value = value if value != 'CHANGED'else ''
                item.setText(_translate("DisplayPanel", str(value), None))

            for i in range(nRow):
                for j in range(nCol):
                    value = matrix[i][j][0] if matrix[i][j][2] != 'CHANGED' else ''

                    value = self.codeUTF8(value)

                    item = QtGui.QTableWidgetItem()
                    tableWidget.setItem(i, j, item)

                    if(i in delRowInfo.keys()):   #分开写四个能在界面上呈现竖能截断红的风格
                        item.setBackgroundColor(red)
                    if(j in addColInfo.keys()):
                        item.setBackgroundColor(blue)
                    if(i in addRowInfo.keys()):
                        item.setBackgroundColor(blue)
                    if(j in delColInfo.keys()):
                        item.setBackgroundColor(red)

                    item.setText(_translate("DisplayPanel", value, None)) #存入数据

            for pos in modifyCellInfo.keys():
                tableWidget.item(pos[0], pos[1]).setBackgroundColor(yellow)  #修改cell黄色效果

        setInGuiTable(self.preMatrix, preMatrix, preRowHead, preColHead)  #初始化展示数据和效果
        setInGuiTable(self.postMatrix, postMatrix, postRowHead, postColHead)

        def setInGuiMessage():
            """
            几行几列发生了修改
            :return:
            """
            strMessage = "共计新增{0}列，删除{1}列".format(len(addColInfo), len(delColInfo))
            self.labelColChange.setText(_translate("DisplayPanel", strMessage, None))
            strMessage = "共计新增{0}行，删除{1}行".format(len(addRowInfo), len(delRowInfo))
            self.labelRowChange.setText(_translate("DisplayPanel", strMessage, None))
            strMessage = "共计{0}个单元格改动".format(len(modifyCellInfo))
            self.labelCellChange.setText(_translate("DisplayPanel", strMessage, None))

        setInGuiMessage()  #底下tab的message 几行几列发生了修改

        def setInGuiTabTable():
            """
            底部信息tab中的table信息载入，并载入信号槽 connect
            :return:
            """
            self.tabColChangeTable.setRowCount(len(delColInfo) + len(addColInfo))
            self.tabRowChangeTable.setRowCount(len(delRowInfo) + len(addRowInfo))
            self.tabCellChangeTable.setRowCount(len(modifyCellInfo))   #初始化3个tab信息

            for index,dict in enumerate(addColInfo.iteritems()):
                addIndex, addNumber = dict
                item = QtGui.QTableWidgetItem()
                self.tabColChangeTable.setItem(index, 0, item)
                item.setText(_translate("DisplayPanel", "新增", None))
                item.setTextColor(deepBlue)

                addLink = "<a href=\"{0}\">{1}</a>".format(addIndex,addNumber)
                hyperlinkLabelAddCol = QtGui.QLabel(addLink,self)

                self.connect(hyperlinkLabelAddCol,QtCore.SIGNAL("linkActivated(QString)"),
                                           self.tabTableColClicked)

                hyperlinkLabelAddCol.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
                self.tabColChangeTable.setCellWidget(index, 1, hyperlinkLabelAddCol)


            for index,dict in enumerate(delColInfo.iteritems()):
                delIndex, delNumber = dict
                item = QtGui.QTableWidgetItem()
                self.tabColChangeTable.setItem(index + len(addColInfo), 0, item)
                item.setText(_translate("DisplayPanel", "删除", None))
                item.setTextColor(red)
                delLink = "<a href=\"{0}\">{1}</a>".format(delIndex,delNumber)
                hyperlinkLabelDelCol = QtGui.QLabel(delLink,self)

                self.connect(hyperlinkLabelDelCol,QtCore.SIGNAL("linkActivated(QString)"),
                                           self.tabTableColClicked)

                hyperlinkLabelDelCol.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
                self.tabColChangeTable.setCellWidget(index + len(addColInfo), 1, hyperlinkLabelDelCol)


            for index,dict in enumerate(addRowInfo.iteritems()):
                addIndex, addNumber = dict
                item = QtGui.QTableWidgetItem()
                self.tabRowChangeTable.setItem(index, 0, item)
                item.setText(_translate("DisplayPanel", "新增", None))
                item.setTextColor(deepBlue)
                addLink = "<a href=\"{0}\">{1}</a>".format(addIndex,addNumber)
                hyperlinkLabelAddRow = QtGui.QLabel(addLink,self)

                self.connect(hyperlinkLabelAddRow,QtCore.SIGNAL("linkActivated(QString)"),
                                           self.tabTableRowClicked)

                hyperlinkLabelAddRow.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
                self.tabRowChangeTable.setCellWidget(index, 1, hyperlinkLabelAddRow)

            for index,dict in enumerate(delRowInfo.iteritems()):
                delIndex, delNumber = dict
                item = QtGui.QTableWidgetItem()
                self.tabRowChangeTable.setItem(index + len(addRowInfo), 0, item)
                item.setText(_translate("DisplayPanel", "删除", None))
                item.setTextColor(red)
                delLink = "<a href=\"{0}\">{1}</a>".format(delIndex,delNumber)
                hyperlinkLabelDelRow = QtGui.QLabel(delLink)

                self.connect(hyperlinkLabelDelRow, QtCore.SIGNAL("linkActivated(QString)"),
                             self.tabTableRowClicked)

                hyperlinkLabelDelRow.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
                self.tabRowChangeTable.setCellWidget(index + len(addRowInfo), 1, hyperlinkLabelDelRow)
                
                
            for index,dict in enumerate(sorted(modifyCellInfo.iteritems())):
                indexCell,value = dict
                rowPre,colPre = value[1]
                rowPost, colPost = value[3]

                strPos = "[{0},{1}],[{2},{3}]".format(rowPre,colPre,rowPost,colPost)
                modifyCellLink = "<a href=\"{0}\">{1}</a>".format(indexCell,strPos)
                hyperlinkLabelChangedCell = QtGui.QLabel(modifyCellLink,self)
                hyperlinkLabelChangedCell.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
                self.tabCellChangeTable.setCellWidget(index, 0, hyperlinkLabelChangedCell)

                self.connect(hyperlinkLabelChangedCell, QtCore.SIGNAL("linkActivated(QString)"),self.tabTableCellClicked)


                item = QtGui.QTableWidgetItem()
                self.tabCellChangeTable.setItem(index, 1, item)
                item.setText(_translate("DisplayPanel", self.codeUTF8(value[0]), None))

                item = QtGui.QTableWidgetItem()
                self.tabCellChangeTable.setItem(index, 2, item)
                item.setText(_translate("DisplayPanel", self.codeUTF8(value[2]), None))


        setInGuiTabTable()   #将数据和链接存入下方tab的table中

        self.tabCellChangeTable.resizeColumnsToContents()
        self.tabCellChangeTable.resizeRowsToContents()
        self.tabColChangeTable.resizeColumnsToContents()
        self.tabColChangeTable.resizeRowsToContents()
        self.tabRowChangeTable.resizeColumnsToContents()
        self.tabRowChangeTable.resizeRowsToContents()    #调整table大小


    def codeUTF8(self,value):
        if isinstance(value, unicode):  # 中文转换编码
            value = value.encode('utf8')
        else:
            value = str(value)
        return value




    """
    调用方法
    """
    def tabTableRowClicked(self,QString):
        """
        行点击联动
        :param QString:
        :return:
        """
        index,boolean = QString.toInt()
        self.matrixTableRowSelected(index)
    def tabTableColClicked(self,QString):
        """
        列点击联动
        :param QString:
        :return:
        """
        #sender = self.sender()
        index,boolean = QString.toInt()
        self.matrixTableColSelected(index)
    def tabTableCellClicked(self,QString):
        """
        改动Cell联动
        :param QString:
        :return:
        """
        pos = eval(str(QString.toUtf8()))
        self.matrixTableCellSelected(pos[0],pos[1])
    def matrixTableClicked(self,item=None):
        """
        展示矩阵选择联动
        :param item:
        :return:
        """
        if item==None:
            return
        self.matrixTableCellSelected(item.row(),item.column())

    def matrixTableCellSelected(self,row,col):
        self.matrixTableColSelected(col)  #先选中某行某列，用于进行跳转，不然不能进行跳转
        self.matrixTableRowSelected(row)
        self.postMatrix.clearSelection()
        self.preMatrix.clearSelection()
        self.postMatrix.item(row,col).setSelected(True)  #最后选中cell
        self.preMatrix.item(row,col).setSelected(True)

    def matrixTableColSelected(self,col):
        """
        列选择在展示中联动
        :param col:
        :return:
        """
        self.postMatrix.clearSelection()
        self.preMatrix.clearSelection()
        self.postMatrix.selectColumn(col)
        self.preMatrix.selectColumn(col)
    def matrixTableRowSelected(self,row):
        """
        行选择展示中联动
        :param row:
        :return:
        """
        self.postMatrix.clearSelection()
        self.preMatrix.clearSelection()
        self.postMatrix.selectRow(row)
        self.preMatrix.selectRow(row)

    def setLabelPath(self,path1,path2):
        self.label.setText(_translate("DisplayPanel", path1, None))
        self.label2.setText(_translate("DisplayPanel", path2, None))




