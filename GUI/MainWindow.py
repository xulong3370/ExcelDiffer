#coding=utf-8
"""
主窗口
"""
from ChooseFile import ChooseFile
from DisplayPanel import DisplayPanel
from AlgorithmThread import AlgorithmThread
from SheetDelOrAddPanel import SheetDelOrAddPanel
from ProgressBar import ProgressBar

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

class MainWindow(object):
    def setupUi(self, MainWindow):
        self.chooseFile = None
        screen = QtGui.QDesktopWidget().screenGeometry()
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(screen.width()/1.5, screen.height()/1.5)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.horizontalLayout = QtGui.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.tabWidget = QtGui.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(_fromUtf8("tabWidget"))
        self.tabWidget.setTabsClosable(True)
        self.tabWidget.tabCloseRequested.connect(self.closeTab)

        self.horizontalLayout.addWidget(self.tabWidget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        MainWindow.setStatusBar(self.statusbar)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 846, 23))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        self.menu = QtGui.QMenu(self.menubar)
        self.menu.setObjectName(_fromUtf8("menu"))
        MainWindow.setMenuBar(self.menubar)
        self.action = QtGui.QAction(MainWindow)
        self.action.setObjectName(_fromUtf8("action"))
        self.actionExit = QtGui.QAction(MainWindow)
        self.actionExit.setObjectName(_fromUtf8("actionExit"))
        self.menu.addAction(self.action)
        self.menu.addAction(self.actionExit)
        self.menubar.addAction(self.menu.menuAction())

        self.action.triggered.connect(self.menuChooseFile)  #选文件
        self.centralwidget.connect(self.actionExit, QtCore.SIGNAL("triggered()"),QtGui.qApp, QtCore.SLOT("quit()"))

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)

        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "ExcelDiff", None))
        #self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Tab 1", None))
        self.menu.setTitle(_translate("MainWindow", "开始", None))
        self.action.setText(_translate("MainWindow", "载入", None))
        self.actionExit.setText(_translate("MainWindow", "退出", None))

    def closeTab(self,index):
        self.tabWidget.removeTab(index)

    def menuChooseFile(self):
        """
        弹出选择文件窗口
        :return:
        """
        self.chooseFile = ChooseFile()
        self.centralwidget.connect(self.chooseFile,QtCore.SIGNAL('executeThread()'),
                                   self.executeAlgorithmInBackground) #链接信号，用于执行后台算法
        self.chooseFile.show()

    def executeAlgorithmInBackground(self):
        """
        执行后台算法
        :return:
        """
        self.initProgressBar()
        self.thread = AlgorithmThread(self.chooseFile.workbook1, self.chooseFile.workbook2,
                                      self.chooseFile.sheets1, self.chooseFile.sheets2)
        self.chooseFile.close()  # 关闭对比栏
        self.centralwidget.connect(self.thread, QtCore.SIGNAL('threadFinished()'),
                                   self.getDataFromThread) #线程获取数据信号
        self.centralwidget.connect(self.thread,QtCore.SIGNAL('threadStopped()'),
                                   self.closeAllProgress) #线程终止信号
        self.centralwidget.connect(self.thread,QtCore.SIGNAL('updateProgressbar(int)'),
                                   self.updateProgressBar) #线程更新进度条
        self.centralwidget.connect(self.thread,QtCore.SIGNAL('updateProgressMessage(QString)'),
                                   self.updateProgressMessage)#线程跟新进度条信息
        self.centralwidget.connect(self.thread,QtCore.SIGNAL('updateProgressTitle(QString,int)'),
                                   self.updateProgressTitle)#线程跟新进度条title信息
        self.centralwidget.connect(self.thread, QtCore.SIGNAL('showError(QString)'), self.showError) #错误显示策略
        self.initProgressBar() #开启进度条
        self.thread.start()   #启动后台线程

    def getDataFromThread(self):
        """
        :param storeList: 信息列表[((fileName1,sheetName1),(fileName2,sheetName2),matrixInfo),....]
        :return:
        """
        count = self.tabWidget.count() #记录当前跳转位置
        tabInfo = self.thread.store  #从thread中获得计算结果
        currentCompareFileTab = QtGui.QTabWidget()
        str = '对比{0}'.format(count + 1)
        self.tabWidget.addTab(currentCompareFileTab,_fromUtf8(str))
        for index,Info in enumerate(tabInfo):
            file1,file2,matrixInfo = Info
            tab = DisplayPanel()
            tab.setResultInGUI(matrixInfo)
            if(file1[1] == file2[1]):
                str = '{0}'.format(self.codeUTF8(file1[1]))
            else:
                str = self.codeUTF8('对比')
            tab.setObjectName(_fromUtf8(str))
            currentCompareFileTab.addTab(tab, _fromUtf8(str))

            strLabel1 = "{0}[{1}]".format(self.codeUTF8(file1[0]),self.codeUTF8(file1[1]))
            strLabel2 = "{0}[{1}]".format(self.codeUTF8(file2[0]),self.codeUTF8(file2[1]))
            tab.setLabelPath(strLabel1,strLabel2)

        if(self.thread.delSheets or self.thread.addSheets):

            sheetDelOrAddTab = SheetDelOrAddPanel()
            currentCompareFileTab.addTab(sheetDelOrAddTab,_fromUtf8('sheet增删'))
            sheetDelOrAddTab.setSheetDelOrAddInfo(self.thread.fileName1,self.thread.fileName2,
                                                  self.thread.addSheets,self.thread.delSheets)

        currentCount = self.tabWidget.count()
        self.tabWidget.setCurrentIndex(currentCount - 1)
        self.cancelProgressBar()  #关闭进度条
        self.currentValue = 0

    def terminateThread(self):
        """
        终止线程
        :return:
        """
        self.reply = QtGui.QMessageBox.question(self.centralwidget, u'警告',u"确定要取消吗？",
                                           QtGui.QMessageBox.Yes,QtGui.QMessageBox.No)
        self.progressBar.hide()
        if self.reply == QtGui.QMessageBox.Yes:
            if(self.thread.isRunning()):
                self.thread.stop()  #为了正常退出程序，立flag，正常退出可能需要结束当前线程。有卡顿
                self.progressBar.show()
                self.progressBar.setDisabled(True) #不能再点击终止
                self.progressBar.updateMessage(u'正在撤销操作，请稍后...')
                self.progressBar.updateWindowTitle(u'终止')
        else:
            if(self.thread.isRunning()):
                self.progressBar.show()
            else:
                self.progressBar.close()

    def closeAllProgress(self):
        """
        表示线程已经安全退出
        :return:
        """
        self.currentValue = 0
        self.progressBar.close()
        self.menu.setDisabled(False)


    def initProgressBar(self):
        """
        开启准备进度条
        :return:
        """
        if (isinstance(self.chooseFile.sheets1,list) and isinstance(self.chooseFile.sheets2,list)):
            modifySheetList = [sheet for sheet in self.chooseFile.sheets1 if sheet in self.chooseFile.sheets2]
            delSheetList = [sheet for sheet in self.chooseFile.sheets1 if sheet not in self.chooseFile.sheets2]
            addSheetList = [sheet for sheet in self.chooseFile.sheets2 if sheet not in self.chooseFile.sheets1]
            sheetsCount = len(modifySheetList) + len(delSheetList) + len(addSheetList)
            max = sheetsCount*100
        else:
            max = 100
        self.progressBar = ProgressBar() #一个进度条
        self.progressBar.updateWindowTitle(u'请稍后...')
        self.currentValue = 0
        self.progressBar.cancelButton.clicked.connect(self.terminateThread)  # 手动终止
        self.progressBar.setRange(0,max)
        self.progressBar.show() #产生一个进度条
        self.menu.setDisabled(True)
        self.progressBar.updateMessage(u'准备中')
        self.progressBar.updateValue(self.currentValue)

    def updateProgressBar(self,progress):
        """
        进度条执行逻辑,执行更新
        :return:
        """
        self.currentValue = progress + self.currentValue
        self.progressBar.updateValue(self.currentValue)
    def updateProgressMessage(self,QString):
        """
        进度条上方label更新
        :param QString:
        :return:
        """
        self.progressBar.updateMessage(QString)
    def updateProgressTitle(self,QString,index):
        """
        进度条上方标题，标明正在处理的多个sheet,第几个进度
        :param QString:
        :return:
        """
        self.progressBar.updateWindowTitle(QString)
        self.currentValue = (index - 1) * 100
        self.progressBar.updateValue(self.currentValue)
    def cancelProgressBar(self):
        """
        关闭进度条逻辑，算法执行完关闭
        :return:
        """
        self.progressBar.close()
        self.menu.setDisabled(False)

    def showError(self,QString):
        """
        主界面产生错误
        :param QString:
        :return:
        """
        reply = QtGui.QMessageBox.warning(self.centralwidget, u'错误',QString, u"确定")
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
