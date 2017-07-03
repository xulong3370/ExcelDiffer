#coding=utf-8
"""
线程后台算法调用,防止卡顿
"""
from Algorithm.DifferAlgorithm import *
from Algorithm.AnalysisMatrixInfo import AnalysisMatrixInfo
from PyQt4 import QtCore,QtGui
from TraceProgress import setGUIMessage

class AlgorithmThread(QtCore.QThread):
    def __init__(self,workbook1,workbook2,sheetList1,sheetList2,parent=None):
        super(AlgorithmThread, self).__init__(parent)
        self.store = []
        self.cancelFlag = False
        self.workbook1 = workbook1
        self.workbook2 = workbook2
        self.sheetList1 = sheetList1
        self.sheetList2 = sheetList2
        fileName1 = self.workbook1.getFilePath()
        fileName2 = self.workbook2.getFilePath()
        self.fileName1 = AlgorithmThread.codeUTF8(fileName1)
        self.fileName2 = AlgorithmThread.codeUTF8(fileName2)
        self.store = []
        self.delSheets = []
        self.addSheets = []

    def run(self):
        try:
            if(isinstance(self.sheetList1,list) and isinstance(self.sheetList2,list)):
                modifySheetList = [sheet for sheet in self.sheetList1 if sheet in self.sheetList2]
                delSheets = [sheet for sheet in self.sheetList1 if sheet not in self.sheetList2]
                addSheets = [sheet for sheet in self.sheetList2 if sheet not in self.sheetList1]
                sheetsCount = len(modifySheetList) + len(delSheets) + len(addSheets)
                index = 1
                for sheet in modifySheetList:
                    if(not self.cancelFlag):
                        titleStrLabel = u"""正在处理第{0}个sheet，共{1}个...""".format(index,sheetsCount)
                        self.emit(QtCore.SIGNAL('updateProgressTitle(QString,int)'),titleStrLabel,index)
                        choose2 = choose1 = AlgorithmThread.codeUTF8(sheet)
                        matrixInfo = self.execAlgorithm(choose1,choose2)
                        self.storeDisplayInfo(self.fileName1,choose1,self.fileName2,choose2,matrixInfo)
                        index = index + 1
                    else:
                        break

                for sheet in delSheets:
                    if(not self.cancelFlag):
                        matrix = DifferAlgorithm.getOriginalMatrix(self.workbook1,sheet)
                        self.delSheets.append((sheet,matrix))
                    else:
                        break
                for sheet in addSheets:
                    if(not self.cancelFlag):
                        matrix = DifferAlgorithm.getOriginalMatrix(self.workbook2, sheet)
                        self.addSheets.append((sheet,matrix))
                    else:
                        break
            else:
                titleStrLabel = u"""正在处理第1个sheet，共1个..."""
                self.emit(QtCore.SIGNAL('updateProgressTitle(QString,int)'), titleStrLabel,1)
                self.sheetList1 = AlgorithmThread.codeUTF8(self.sheetList1)  #转码
                self.sheetList2 = AlgorithmThread.codeUTF8(self.sheetList2)
                matrixInfo = self.execAlgorithm(self.sheetList1,self.sheetList2)
                self.storeDisplayInfo(self.fileName1, self.sheetList1, self.fileName2, self.sheetList2, matrixInfo)

            if(not self.cancelFlag):
                self.emit(QtCore.SIGNAL('threadFinished()'))  # 完成后发送数据到主窗体
        except MemoryError,e:
            self.emit(QtCore.SIGNAL('showError(QString)'), e.__class__.__name__)
        finally:
            self.emit(QtCore.SIGNAL('threadStopped()')) #表示线程安全退出

    def stop(self):
        self.cancelFlag = True

    def storeDisplayInfo(self,fileName1,sheetName1,fileName2,sheetName2,matrixInfo):
        """
        存储展示信息
        :return:
        """
        self.store.append(((fileName1,sheetName1),(fileName2,sheetName2),matrixInfo))
    def execAlgorithm(self,choose1,choose2):
        """
        执行算法
        :param choose1:
        :param choose2:
        :return:
        """
        setGUIMessage(self)
        preMatrixByRow, postMatrixByRow, preMatrixByCol, postMatrixByCol,preNRow,preNCol,postNRow,postNCol \
            = DifferAlgorithm.getDifferCellMatrix(self.workbook1, self.workbook2,choose1, choose2)
        matrixInfo = AnalysisMatrixInfo.analysis(preMatrixByRow, postMatrixByRow, preMatrixByCol, postMatrixByCol,
                                                 preNRow,preNCol,postNRow,postNCol)
        return matrixInfo

    @staticmethod
    def codeUTF8(s):
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